# config.py
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# プロジェクトルート基準の絶対パス解決
BASE_DIR = Path(__file__).resolve().parent

# パスの外部化（環境変数未定義時はデフォルト値を適用）
INPUT_DIR = BASE_DIR / os.getenv("INPUT_DIR", "inputs")

# ユーザーの「ドキュメント/日報」ルートディレクトリを取得
BASE_DOCUMENTS_DIR = Path.home() / "Documents" / "日報"
BASE_DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)

# 従来の互換性を保つ場合、またはデフォルトの出力先とする場合
OUTPUT_DIR = BASE_DOCUMENTS_DIR

TEMPLATE_PATH = BASE_DIR / os.getenv("TEMPLATE_PATH", "inputs/aspark_template.pptx")

# 起動時のディレクトリ自動生成（原則5の連動）
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)