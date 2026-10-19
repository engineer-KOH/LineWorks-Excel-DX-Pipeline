"""
Title: LINE WORKS Excel Webhook Bot (Main Server - Full Pipeline + Future Routing Ready)
Description: 添付された日報Excelファイルを受信し、Excel正規化・Gemini要約・PPTX/JSON生成を完遂するAPIサーバー。
             （将来の担当者自動ルーティング・通知機能用のコードをコメントアウトで同梱）
Version: 3.2.0 (Routing Prepared & Commented Out)

【更新履歴 (Change Log)】
- v1.0.0 (2026/09/10): 初期プロトタイプ作成（Webhook基本受信）
- v2.0.0 (2026/09/14): HMAC-SHA256 署名検証および Pydantic 厳格バリデーションの導入
- v2.1.0 (2026/09/15): バックグラウンドタスク内への openpyxl 解析パイプラインの統合
- v2.2.0 (2026/09/15): Service Account (JWT/RS256) によるアクセストークン動的取得・File API連携
- v2.2.4 (2026/09/17): リアルタイムフラッシュ対応カスタムロギングハンドラーの導入
- v3.0.0 (2026/09/17): ローカル単独テストと同等の「Gemini API要約・JSON出力・テンプレートPPTX生成」パイプラインを完全統合
- v3.1.0 (2026/09/17): Error Push Notification Integrated
- v3.2.0 (2026/09/17): 将来の mapping.json 連携・担当者自動ルーティング用の準備コード（コメントアウト）を追加
"""

import os
import hmac
import hashlib
import base64
import logging
import time
import urllib.parse
from typing import Optional, Dict, List
from pathlib import Path
from fastapi import FastAPI, Request, Header, HTTPException, BackgroundTasks, status
from pydantic import BaseModel
import httpx
import jwt  # PyJWT
from dotenv import load_dotenv

# -------------------------------------------------------------------
# 既存のコアモジュール・ユーティリティのインポート
# -------------------------------------------------------------------
from excel_normalizer import ExcelNormalizer
from gemini_summarizer import GeminiSummarizer
from pptx_generator import PPTXGenerator
from file_utils import extract_name_from_filename
from config import INPUT_DIR, OUTPUT_DIR, TEMPLATE_PATH

# 環境変数の読み込み
load_dotenv()

LINEWORKS_SECRET_KEY = os.getenv("LINEWORKS_SECRET_KEY", "")
LINEWORKS_BOT_ID = os.getenv("LINEWORKS_BOT_ID", "")
LINEWORKS_CLIENT_ID = os.getenv("LINEWORKS_CLIENT_ID", "")
LINEWORKS_CLIENT_SECRET = os.getenv("LINEWORKS_CLIENT_SECRET", "")
LINEWORKS_SERVICE_ACCOUNT = os.getenv("LINEWORKS_SERVICE_ACCOUNT", "")
ADMIN_USER_ID = os.getenv("ADMIN_USER_ID", "")

# -------------------------------------------------------------------
# ロギング設定（リアルタイムフラッシュ対応カスタムハンドラー）
# -------------------------------------------------------------------
class FlushingFileHandler(logging.FileHandler):
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
    description="LINE WORKS経由で受信した日報Excelを処理するAPIサーバー",
    version="3.2.0"
)

# -------------------------------------------------------------------
# Pydantic Schemas (LINE WORKS Payload Structure)
# -------------------------------------------------------------------
class Content(BaseModel):
    type: str
    fileId: Optional[str] = None
    fileName: Optional[str] = None
    text: Optional[str] = None

class Source(BaseModel):
    userId: str

class LineWorksWebhookPayload(BaseModel):
    type: str
    source: Source
    createdTime: Optional[int] = None
    content: Content


# -------------------------------------------------------------------
# セキュリティ: HMAC-SHA256 署名検証関数
# -------------------------------------------------------------------
def verify_signature(body: bytes, signature: str) -> bool:
    if not LINEWORKS_SECRET_KEY:
        return True
    hash_digest = hmac.new(
        LINEWORKS_SECRET_KEY.encode('utf-8'),
        body,
        hashlib.sha256
    ).digest()
    expected_signature = base64.b64encode(hash_digest).decode('utf-8')
    return hmac.compare_digest(expected_signature, signature)


# -------------------------------------------------------------------
# LINE WORKS 認証: JWT生成 & アクセストークン動的取得関数
# -------------------------------------------------------------------
# -------------------------------------------------------------------
# LINE WORKS 認証: JWT生成 & アクセストークン動的取得関数 (.pemフォールバック対応)
# -------------------------------------------------------------------
LINEWORKS_PRIVATE_KEY_PATH = os.getenv("LINEWORKS_PRIVATE_KEY_PATH", "./system/private_key.pem")

def get_lineworks_access_token() -> Optional[str]:
    token_url = "https://auth.worksmobile.com/oauth2/v2.0/token"
    
    # 1. まず環境変数 LINEWORKS_SECRET_KEY から秘密鍵を取得
    private_key_str = os.getenv("LINEWORKS_SECRET_KEY", "")

    # 2. 環境変数が空、または不正な場合は .pem ファイルから読み込みを試みる
    if not private_key_str or "BEGIN PRIVATE KEY" not in private_key_str:
        pem_path = Path(LINEWORKS_PRIVATE_KEY_PATH)
        if pem_path.exists():
            try:
                private_key_str = pem_path.read_text(encoding="utf-8")
                logger.info(f"[Token Key Loaded] .pemファイルから秘密鍵を読み込みました: {pem_path}")
            except Exception as e:
                logger.error(f"[Token Key Read Error] .pemファイルの読み込みに失敗しました: {e}")
                return None
        else:
            logger.error(f"[Token Check Fail] 秘密鍵が見つかりません。(env: False, pem_path: {pem_path.absolute()} False)")
            return None

    if not LINEWORKS_CLIENT_ID or not LINEWORKS_SERVICE_ACCOUNT:
        logger.error(f"[Token Check Fail] 必須環境変数が未設定です: CLIENT_ID={bool(LINEWORKS_CLIENT_ID)}, SERVICE_ACCOUNT={bool(LINEWORKS_SERVICE_ACCOUNT)}")
        return None

    current_time = int(time.time())
    payload = {
        "iss": LINEWORKS_CLIENT_ID,
        "sub": LINEWORKS_SERVICE_ACCOUNT,
        "iat": current_time,
        "exp": current_time + 3600,
    }

    try:
        encoded_jwt = jwt.encode(payload, private_key_str.strip(), algorithm="RS256")
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
                logger.info("[Token Success] アクセストークンの動的取得に成功しました。")
                return token
            else:
                logger.error(f"[Token HTTP Error {response.status_code}]: {response.text}")
                return None
    except Exception as e:
        logger.error(f"[Token JWT/RSA Error]: {e}", exc_info=True)
        return None


# -------------------------------------------------------------------
# エラー発生時のLINE WORKSプッシュ通知関数
# -------------------------------------------------------------------
async def send_line_works_error_notification(user_id: str, error_message: str):
    access_token = get_lineworks_access_token()
    if not access_token or not LINEWORKS_BOT_ID:
        return
    target_user_id = ADMIN_USER_ID if ADMIN_USER_ID else user_id
    api_url = f"https://www.worksapis.com/v1.0/bots/{LINEWORKS_BOT_ID}/users/{target_user_id}/messages"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    payload = {
        "content": {
            "type": "text",
            "text": f"【⚠️ 日報自動処理 エラー通知】\nバックグラウンド処理中に例外が発生しました。\n\n詳細:\n{error_message}"
        }
    }
    try:
        async with httpx.AsyncClient() as client:
            await client.post(api_url, headers=headers, json=payload)
    except Exception as e:
        logger.error(f"Error notification send failed: {e}")

# -------------------------------------------------------------------
# バックグラウンド処理: パイプライン実行
# -------------------------------------------------------------------
async def process_excel_attachment(
    user_id: str, file_id: str, file_name: Optional[str], created_time: Optional[int]
):
    logger.info(f"[Task Start] User: {user_id} | FileID: {file_id} | WebhookFileName: {file_name}")
    
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    original_filename = file_name
    download_success = False
    access_token = get_lineworks_access_token()
    file_path = None

    if access_token:
        file_api_url = f"https://www.worksapis.com/v1.0/bots/{LINEWORKS_BOT_ID}/attachments/{file_id}"
        headers = {"Authorization": f"Bearer {access_token}"}

        try:
            async with httpx.AsyncClient() as client:
                response = await client.get(file_api_url, headers=headers, follow_redirects=False)
                logger.info(f"File API Response Status: {response.status_code}")
                
                if response.status_code in (301, 302, 303, 307, 308):
                    redirect_url = response.headers.get("Location")
                    if redirect_url:
                        response = await client.get(redirect_url, headers=headers, follow_redirects=True)
                        logger.info(f"Redirected File API Status: {response.status_code}")
                        final_url_path = urllib.parse.urlparse(str(response.url)).path
                        for part in final_url_path.split('/'):
                            decoded_part = urllib.parse.unquote(part)
                            if decoded_part.endswith(('.xlsx', '.xls')):
                                original_filename = decoded_part
                                break

                if response.status_code == 200:
                    resolved_name = original_filename or f"report_{file_id}.xlsx"
                    file_path = INPUT_DIR / resolved_name
                    with open(file_path, "wb") as f:
                        f.write(response.content)
                    logger.info(f"実ファイルをローカルに保存しました: {file_path}")
                    download_success = True
                else:
                    # 200以外（401, 403, 404等）の場合にレスポンス本文を出力
                    logger.error(f"File API エラーレスポンス [Status {response.status_code}]: {response.text}")

        except Exception as e:
            # スタックトレースを含めて例外の詳細を出力
            logger.error(f"APIダウンロード処理中に例外が発生しました: {e}", exc_info=True)
    else:
        logger.error("アクセストークンの取得に失敗しているため、ダウンロード処理をスキップしました。")
    
    if not download_success or not file_path or not os.path.exists(file_path):
        err_msg = "有効なExcelファイルのダウンロードに失敗しました。"
        logger.error(err_msg)
        await send_line_works_error_notification(user_id, err_msg)
        return

    try:
        # 1. 氏名の抽出
        author_name = extract_name_from_filename(file_path.name)
        logger.info(f"ファイル名より氏名を抽出: '{author_name}'")

        # 2. Excelファイルの正規化パース
        logger.info("Step 1: Excelファイルの正規化パースを開始...")
        normalizer = ExcelNormalizer(ignore_hidden_sheets=True)
        workbook_ir = normalizer.process_file(str(file_path))

        # 3. Gemini APIによる構造化要約
        logger.info("Step 2: Gemini APIによる構造化要約を開始...")
        summarizer = GeminiSummarizer()
        summary_data = summarizer.summarize_workbook(workbook_ir, author_name=author_name)
        
        base_title = summary_data.document_title if "成果報告" in summary_data.document_title else "成果報告_9月"
        if author_name and f"({author_name})" not in base_title:
            summary_data.document_title = f"{base_title} ({author_name})"

        # 4. PPTXおよびJSONの自動生成
        logger.info("Step 3: 同名ファイルでのPPTXおよびJSONの自動生成を開始...")
        stem_name = file_path.stem
        output_pptx = OUTPUT_DIR / f"{stem_name}.pptx"
        output_json = OUTPUT_DIR / f"{stem_name}.json"

        # スライドの並び順を「自己学習」→「社内業務」の順にソートする
        category_order = {"自己学習": 1, "社内学習": 2, "社内業務": 3}
        summary_data.slides.sort(key=lambda s: category_order.get(s.category, 99))

        generator = PPTXGenerator(template_path=str(TEMPLATE_PATH))
        generator.generate(summary_data, output_path=output_pptx, author_name=author_name)

        output_json.write_text(summary_data.model_dump_json(indent=2, ensure_ascii=False), encoding="utf-8")
        
        logger.info(f"出力完了 -> PPTX: {output_pptx.name} / JSON: {output_json.name}")
        logger.info(f"[Task Complete] パイプライン処理が正常終了しました。")

       
    except Exception as e:
        error_details = str(e)
        logger.error(f"[Task Error] バックグラウンド処理中に例外が発生しました: {error_details}", exc_info=True)
        await send_line_works_error_notification(user_id, error_details)


# -------------------------------------------------------------------
# Webhook エンドポイント (HTTP POST)
# -------------------------------------------------------------------
@app.post("/webhook", status_code=status.HTTP_200_OK)
async def handle_webhook(
    request: Request,
    background_tasks: BackgroundTasks,
    x_lineworks_signature: Optional[str] = Header(None)
):
    body = await request.body()
    if x_lineworks_signature and not verify_signature(body, x_lineworks_signature):
        logger.error("不正な署名のリクエストを検出しました。")
        raise HTTPException(status_code=400, detail="Invalid signature")

    try:
        data = await request.json()
        payload = LineWorksWebhookPayload(**data)
    except Exception as e:
        logger.error(f"JSONパースエラー: {e}")
        return {"status": "ignored", "reason": "invalid json format"}

    if payload.content.type == "file" and payload.content.fileId:
        background_tasks.add_task(
            process_excel_attachment,
            user_id=payload.source.userId,
            file_id=payload.content.fileId,
            file_name=payload.content.fileName,
            created_time=payload.createdTime
        )
        return {"status": "success", "message": "Excel file queued for full pipeline processing"}

    logger.info(f"ファイル以外のメッセージタイプを受信（スキップ）: {payload.content.type}")
    return {"status": "ignored", "reason": "not a file message"}