"""Phase 3 integrated automated generation pipeline script."""

import logging
from pathlib import Path
import sys

from config import INPUT_DIR, OUTPUT_DIR, TEMPLATE_PATH
from excel_normalizer import ExcelNormalizer
from file_utils import extract_name_from_filename
from gemini_summarizer import GeminiSummarizer
from pptx_generator import PPTXGenerator

# Logging setup (console output + file output to logs/gemini_debug.log)
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
log_path = log_dir / "gemini_debug.log"

root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO)

has_file_handler = any(
    isinstance(h, logging.FileHandler) for h in root_logger.handlers
)
if not has_file_handler:
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    )
    root_logger.addHandler(file_handler)

logger = logging.getLogger(__name__)


def main():
    """Execute the phase 3 integrated automated generation pipeline."""
    logger.info("=== Phase 3 Integrated Automated Pipeline Started ===")

    excel_path = INPUT_DIR / "成果報告日報_氏名_2026_09_04-09_30.xlsx"

    if not excel_path.exists():
        logger.error(f"Input Excel file not found: {excel_path.absolute()}")
        sys.exit(1)

    author_name = extract_name_from_filename(excel_path.name)
    logger.info(f"Extracted author name from filename: '{author_name}'")

    logger.info("Step 1: Starting Excel normalization parsing...")
    normalizer = ExcelNormalizer(ignore_hidden_sheets=True)
    workbook_ir = normalizer.process_file(str(excel_path))

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
    else:
        summary_data.document_title = base_title

    logger.info("Step 3: Starting automatic generation of PPTX and JSON...")
    output_pptx = OUTPUT_DIR / f"{excel_path.stem}.pptx"
    output_json = OUTPUT_DIR / f"{excel_path.stem}.json"

    generator = PPTXGenerator(template_path=str(TEMPLATE_PATH))
    generator.generate(
        summary_data, output_path=output_pptx, author_name=author_name
    )

    output_json.write_text(
        summary_data.model_dump_json(indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    logger.info(f"JSON save completed: {output_json.absolute()}")

    logger.info("=== All pipeline processes completed successfully ===")


if __name__ == "__main__":
    main()