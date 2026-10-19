"""
==============================================================================
スクリプト名 : run_phase3_pipeline.py
機能概要     : Excel読み込みからGemini API要約、PPTX/JSON生成までの一連処理を自動化
               （表紙タイトル固定・パイプライン完全版）
作成日       : 2026-09-10
更新履歴     : 2026-09-11 v3.4.0 表紙タイトル生成ルール補正対応
            : 2026-09-17 パスの外部化対応
                          - ハードコードされていた入出力パス、テンプレートパスを
                          - `config.py` および `.env` 経由で取得するよう変更。
                          - 異なる実行環境やフォルダ構成への移行性を向上。
            : 2026-09-17 ログ出力強化対応
                          - デバッグ用ログを `logs/gemini_debug.log`（UTF-8）へ
                          - 自動保存するようロギング設定を拡張。
==============================================================================
"""

import sys
import logging
from pathlib import Path

# config から外部化されたパスをインポート
from config import INPUT_DIR, OUTPUT_DIR, TEMPLATE_PATH
from excel_normalizer import ExcelNormalizer
from gemini_summarizer import GeminiSummarizer
from pptx_generator import PPTXGenerator
from file_utils import extract_name_from_filename

# ------------------------------------------------------------------------------
# ロギング設定（コンソール出力 ＋ logs/gemini_debug.log へのファイル出力）
# ------------------------------------------------------------------------------
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
log_path = log_dir / "gemini_debug.log"

# ルートロガーを取得
root_logger = logging.getLogger()
root_logger.setLevel(logging.INFO)

# すでに FileHandler が追加されていない場合のみ、安全に追加する（重複防止）
has_file_handler = any(isinstance(h, logging.FileHandler) for h in root_logger.handlers)
if not has_file_handler:
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    root_logger.addHandler(file_handler)

logger = logging.getLogger(__name__)

def main():
    logger.info("=== フェーズ3 統合自動生成パイプライン 開始 ===")

    # INPUT_DIR を基準にファイルパスを構築
    excel_path = INPUT_DIR / "成果報告日報_氏名_2026_09_04-09_30.xlsx"

    if not excel_path.exists():
        logger.error(f"入力Excelファイルが見つかりません: {excel_path.absolute()}")
        sys.exit(1)

    author_name = extract_name_from_filename(excel_path.name)
    logger.info(f"ファイル名より氏名を抽出: '{author_name}'")

    logger.info("Step 1: Excelファイルの正規化パースを開始...")
    normalizer = ExcelNormalizer(ignore_hidden_sheets=True)
    workbook_ir = normalizer.process_file(str(excel_path))

    logger.info("Step 2: Gemini APIによる構造化要約を開始...")
    summarizer = GeminiSummarizer()
    summary_data = summarizer.summarize_workbook(workbook_ir, author_name=author_name)
    
    base_title = summary_data.document_title if "成果報告" in summary_data.document_title else "成果報告_9月"
    if author_name and f"({author_name})" not in base_title:
        summary_data.document_title = f"{base_title} ({author_name})"
    else:
        summary_data.document_title = base_title

    logger.info("Step 3: PPTXおよびJSONの自動生成を開始...")
    # TEMPLATE_PATH、OUTPUT_DIR を config 経由で渡す
    output_pptx = OUTPUT_DIR / f"{excel_path.stem}.pptx"
    output_json = OUTPUT_DIR / f"{excel_path.stem}.json"

    generator = PPTXGenerator(template_path=str(TEMPLATE_PATH))
    generator.generate(summary_data, output_path=output_pptx, author_name=author_name)

    output_json.write_text(summary_data.model_dump_json(indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"JSON保存完了: {output_json.absolute()}")

    logger.info("=== パイプライン全工程が正常完了しました ===")

if __name__ == "__main__":
    main()