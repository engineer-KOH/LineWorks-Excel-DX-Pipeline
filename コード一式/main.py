"""LINE WORKS Excel Webhook Bot (Main Server - Full Pipeline + Future Routing Ready).

Description: API server that receives attached report Excel files, executes normalization,
             Gemini summarization, and generates PPTX/JSON.
Version: 3.2.0 (Routing Prepared & Commented Out)
"""

# Standard library imports
import base64
import hashlib
import hmac
import logging
import os
from pathlib import Path
import time
import typing
import urllib.parse

# Third-party library imports
from dotenv import load_dotenv
from fastapi import BackgroundTasks, FastAPI, Header, HTTPException, Request, status
import httpx
import jwt  # PyJWT
from pydantic import BaseModel

# Local module imports
from config import INPUT_DIR, OUTPUT_DIR, TEMPLATE_PATH
from excel_normalizer import ExcelNormalizer
from file_utils import extract_name_from_filename
from gemini_summarizer import GeminiSummarizer
from pptx_generator import PPTXGenerator

# Load environment variables
load_dotenv()

LINEWORKS_SECRET_KEY = os.getenv("LINEWORKS_SECRET_KEY", "")
LINEWORKS_BOT_ID = os.getenv("LINEWORKS_BOT_ID", "")
LINEWORKS_CLIENT_ID = os.getenv("LINEWORKS_CLIENT_ID", "")
LINEWORKS_CLIENT_SECRET = os.getenv("LINEWORKS_CLIENT_SECRET", "")
LINEWORKS_SERVICE_ACCOUNT = os.getenv("LINEWORKS_SERVICE_ACCOUNT", "")
ADMIN_USER_ID = os.getenv("ADMIN_USER_ID", "")


# -------------------------------------------------------------------
# Logging Setup (Custom handler supporting real-time flushing)
# -------------------------------------------------------------------
class FlushingFileHandler(logging.FileHandler):
    """File handler that flushes log records immediately."""

    def emit(self, record):
        super().emit(record)
        self.flush()


log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
log_path = log_dir / "gemini_debug.log"

logger = logging.getLogger("lineworks_bot")
logger.setLevel(logging.INFO)

if logger.hasHandlers():
    logger.handlers.clear()

log_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

file_handler = FlushingFileHandler(log_path, encoding="utf-8")
file_handler.setFormatter(log_format)
logger.addHandler(file_handler)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(log_format)
logger.addHandler(stream_handler)

app = FastAPI(
    title="LINE WORKS Excel Webhook Bot (Routing Prepared)",
    description="API server to process report Excel files received via LINE WORKS.",
    version="3.2.0",
)


# -------------------------------------------------------------------
# Pydantic Schemas (LINE WORKS Payload Structure)
# -------------------------------------------------------------------
class Content(BaseModel):
    """Schema for message content details."""

    type: str
    fileId: typing.Optional[str] = None
    fileName: typing.Optional[str] = None
    text: typing.Optional[str] = None


class Source(BaseModel):
    """Schema for event source details."""

    userId: str


class LineWorksWebhookPayload(BaseModel):
    """Schema for LINE WORKS incoming webhook payload."""

    type: str
    source: Source
    createdTime: typing.Optional[int] = None
    content: Content


# -------------------------------------------------------------------
# Security: HMAC-SHA256 Signature Verification
# -------------------------------------------------------------------
def verify_signature(body: bytes, signature: str) -> bool:
    """Verify HMAC-SHA256 signature from LINE WORKS webhook headers."""
    if not LINEWORKS_SECRET_KEY:
        return True
    hash_digest = hmac.new(
        LINEWORKS_SECRET_KEY.encode("utf-8"), body, hashlib.sha256
    ).digest()
    expected_signature = base64.b64encode(hash_digest).decode("utf-8")
    return hmac.compare_digest(expected_signature, signature)


# -------------------------------------------------------------------
# LINE WORKS Authentication: JWT Generation & Access Token Dynamic Retrieval
# -------------------------------------------------------------------
LINEWORKS_PRIVATE_KEY_PATH = os.getenv(
    "LINEWORKS_PRIVATE_KEY_PATH", "./system/private_key.pem"
)


def get_lineworks_access_token() -> typing.Optional[str]:
    """Generate JWT, authenticate with LINE WORKS, and retrieve an access token."""
    token_url = "https://auth.worksmobile.com/oauth2/v2.0/token"

    # 1. Retrieve private key string from environment variable
    private_key_str = os.getenv("LINEWORKS_SECRET_KEY", "")

    # 2. Fall back to loading from .pem file if environment variable is empty or invalid
    if not private_key_str or "BEGIN PRIVATE KEY" not in private_key_str:
        pem_path = Path(LINEWORKS_PRIVATE_KEY_PATH)
        if pem_path.exists():
            try:
                private_key_str = pem_path.read_text(encoding="utf-8")
                logger.info(
                    f"[Token Key Loaded] Loaded private key from .pem file: {pem_path}"
                )
            except Exception as e:
                logger.error(
                    f"[Token Key Read Error] Failed to read .pem file: {e}"
                )
                return None
        else:
            logger.error(
                f"[Token Check Fail] Private key not found (env: False, pem_path: {pem_path.absolute()} False)"
            )
            return None

    if not LINEWORKS_CLIENT_ID or not LINEWORKS_SERVICE_ACCOUNT:
        logger.error(
            f"[Token Check Fail] Missing required env vars: CLIENT_ID={bool(LINEWORKS_CLIENT_ID)}, SERVICE_ACCOUNT={bool(LINEWORKS_SERVICE_ACCOUNT)}"
        )
        return None

    current_time = int(time.time())
    payload = {
        "iss": LINEWORKS_CLIENT_ID,
        "sub": LINEWORKS_SERVICE_ACCOUNT,
        "iat": current_time,
        "exp": current_time + 3600,
    }

    try:
        encoded_jwt = jwt.encode(
            payload, private_key_str.strip(), algorithm="RS256"
        )
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": encoded_jwt,
            "client_id": LINEWORKS_CLIENT_ID,
            "client_secret": LINEWORKS_CLIENT_SECRET,
            "scope": "bot",
        }
        with httpx.Client() as client:
            response = client.post(token_url, headers=headers, data=data)
            if response.status_code == 200:
                token = response.json().get("access_token")
                logger.info(
                    "[Token Success] Successfully retrieved access token."
                )
                return token
            else:
                logger.error(
                    f"[Token HTTP Error {response.status_code}]: {response.text}"
                )
                return None
    except Exception as e:
        logger.error(f"[Token JWT/RSA Error]: {e}", exc_info=True)
        return None


# -------------------------------------------------------------------
# Push Error Notification to LINE WORKS
# -------------------------------------------------------------------
async def send_line_works_error_notification(user_id: str, error_message: str):
    """Send an error notification message to the user or admin via LINE WORKS bot API."""
    access_token = get_lineworks_access_token()
    if not access_token or not LINEWORKS_BOT_ID:
        return
    target_user_id = ADMIN_USER_ID if ADMIN_USER_ID else user_id
    api_url = f"https://www.worksapis.com/v1.0/bots/{LINEWORKS_BOT_ID}/users/{target_user_id}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    payload = {
        "content": {
            "type": "text",
            "text": f"【⚠️ 日報自動処理 エラー通知】\nバックグラウンド処理中に例外が発生しました。\n\n詳細:\n{error_message}",
        }
    }
    try:
        async with httpx.AsyncClient() as client:
            await client.post(api_url, headers=headers, json=payload)
    except Exception as e:
        logger.error(f"Error notification send failed: {e}")


# -------------------------------------------------------------------
# Background Task: Pipeline Execution
# -------------------------------------------------------------------
async def process_excel_attachment(
    user_id: str,
    file_id: str,
    file_name: typing.Optional[str],
    created_time: typing.Optional[int],
):
    """Process downloaded Excel attachment through the normalization, summarization, and generation pipeline."""
    logger.info(
        f"[Task Start] User: {user_id} | FileID: {file_id} | WebhookFileName: {file_name}"
    )

    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    original_filename = file_name
    download_success = False
    access_token = get_lineworks_access_token()
    file_path = None

    if access_token:
        file_api_url = (
            f"https://www.worksapis.com/v1.0/bots/{LINEWORKS_BOT_ID}/attachments/{file_id}"
        )
        headers = {"Authorization": f"Bearer {access_token}"}

        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(
                    file_api_url, headers=headers, follow_redirects=False
                )
                logger.info(
                    f"File API Response Status: {response.status_code}"
                )

                if response.status_code in (301, 302, 303, 307, 308):
                    redirect_url = response.headers.get("Location")
                    if redirect_url:
                        response = await client.get(
                            redirect_url, headers=headers, follow_redirects=True
                        )
                        logger.info(
                            f"Redirected File API Status: {response.status_code}"
                        )
                        final_url_path = urllib.parse.urlparse(
                            str(response.url)
                        ).path
                        for part in final_url_path.split("/"):
                            decoded_part = urllib.parse.unquote(part)
                            if decoded_part.endswith((".xlsx", ".xls")):
                                original_filename = decoded_part
                                break

                if response.status_code == 200:
                    resolved_name = (
                        original_filename or f"report_{file_id}.xlsx"
                    )
                    file_path = INPUT_DIR / resolved_name
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    logger.info(
                        f"Saved file locally to destination: {file_path}"
                    )
                    download_success = True
                else:
                    logger.error(
                        f"File API Error Response [Status {response.status_code}]: {response.text}"
                    )

        except Exception as e:
            logger.error(
                f"Exception occurred during file API download: {e}",
                exc_info=True,
            )
    else:
        logger.error(
            "Failed to retrieve access token; skipping download process."
        )

    if not download_success or not file_path or not os.path.exists(file_path):
        err_msg = "Failed to download a valid Excel file."
        logger.error(err_msg)
        await send_line_works_error_notification(user_id, err_msg)
        return

    try:
        # 1. Extract author name from filename
        author_name = extract_name_from_filename(file_path.name)
        logger.info(f"Extracted author name from filename: '{author_name}'")

        # 2. Normalize and parse Excel file
        logger.info("Step 1: Starting Excel normalization parsing...")
        normalizer = ExcelNormalizer(ignore_hidden_sheets=True)
        workbook_ir = normalizer.process_file(str(file_path))

        # 3. Generate structured summary using Gemini API
        logger.info("Step 2: Starting Gemini API structured summarization...")
        summarizer = GeminiSummarizer()
        summary_data = summarizer.summarize_workbook(
            workbook_ir, author_name=author_name
        )

        base_title = (
            summary_data.document_title
            if "成果報告" in summary_data.document_title
            else "成果報告_9月"
        )
        if author_name and f"({author_name})" not in base_title:
            summary_data.document_title = f"{base_title} ({author_name})"

        # 4. Automatically generate PPTX and JSON files
        logger.info(
            "Step 3: Starting automatic generation of PPTX and JSON files..."
        )
        stem_name = file_path.stem
        output_pptx = OUTPUT_DIR / f"{stem_name}.pptx"
        output_json = OUTPUT_DIR / f"{stem_name}.json"

        # Sort slides by category order: Self-learning -> Internal operations
        category_order = {"自己学習": 1, "社内学習": 2, "社内業務": 3}
        summary_data.slides.sort(
            key=lambda s: category_order.get(s.category, 99)
        )

        generator = PPTXGenerator(template_path=str(TEMPLATE_PATH))
        generator.generate(
            summary_data, output_path=output_pptx, author_name=author_name
        )

        output_json.write_text(
            summary_data.model_dump_json(indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        logger.info(
            f"Export completed -> PPTX: {output_pptx.name} / JSON: {output_json.name}"
        )
        logger.info("[Task Complete] Pipeline processing finished successfully.")

    except Exception as e:
        error_details = str(e)
        logger.error(
            f"[Task Error] Exception occurred during background processing: {error_details}",
            exc_info=True,
        )
        await send_line_works_error_notification(user_id, error_details)


# -------------------------------------------------------------------
# Webhook Endpoint (HTTP POST)
# -------------------------------------------------------------------
@app.post("/webhook", status_code=status.HTTP_200_OK)
async def handle_webhook(
    request: Request,
    background_tasks: BackgroundTasks,
    x_lineworks_signature: typing.Optional[str] = Header(None),
):
    """Receive incoming webhook events from LINE WORKS and trigger background processing if applicable."""
    body = await request.body()
    if x_lineworks_signature and not verify_signature(
        body, x_lineworks_signature
    ):
        logger.error("Detected request with invalid signature.")
        raise HTTPException(status_code=400, detail="Invalid signature")

    try:
        data = await request.json()
        payload = LineWorksWebhookPayload(**data)
    except Exception as e:
        logger.error(f"JSON parsing error: {e}")
        return {"status": "ignored", "reason": "invalid json format"}

    if payload.content.type == "file" and payload.content.fileId:
        background_tasks.add_task(
            process_excel_attachment,
            user_id=payload.source.userId,
            file_id=payload.content.fileId,
            file_name=payload.content.fileName,
            created_time=payload.createdTime,
        )
        return {
            "status": "success",
            "message": "Excel file queued for full pipeline processing",
        }

    logger.info(
        f"Received non-file message type (skipped): {payload.content.type}"
    )
    return {"status": "ignored", "reason": "not a file message"}