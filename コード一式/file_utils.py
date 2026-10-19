"""
==============================================================================
モジュール名 : file_utils.py
機能概要     : ファイル名のサニタイズ、同名回避（連番付与）、氏名抽出ユーティリティ
作成日       : 2026-09-10
更新履歴     :
  - 2026-09-10: 初版作成
  - 2026-09-11: 重複していた import re の記述を整理・クリーンアップ。
==============================================================================
"""
import re
from datetime import datetime
from pathlib import Path


def sanitize_filename(filename: str) -> str:
    """
    Windows/Linuxで使用不可な記号を安全な文字に置換する
    """
    invalid_chars = r'[\\/:*?"<>|]'
    return re.sub(invalid_chars, "_", filename)


def get_unique_filepath(output_dir: Path, base_name: str, extension: str = ".pptx", add_date: bool = True) -> Path:
    """
    サニタイズ済みの安全なファイル名を生成し、同名ファイルが存在する場合は連番を付与する
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    clean_name = sanitize_filename(base_name)

    if add_date:
        today_str = datetime.now().strftime("%Y%m%d")
        clean_name = f"{clean_name}_{today_str}"

    target_path = output_dir / f"{clean_name}{extension}"

    counter = 1
    while target_path.exists():
        target_path = output_dir / f"{clean_name}({counter}){extension}"
        counter += 1

    return target_path


def extract_name_from_filename(filename: str) -> str:
    """
    ファイル名から氏名を抽出する
    例: "成果報告日報_末永_2026_09_04-09_30.xlsx" -> "末永"
    """
    parts = Path(filename).stem.split("_")
    if len(parts) >= 2:
        return parts[1]
    
    match = re.search(r"[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]+", filename)
    if match:
        return match.group(0)
        
    return "担当者"

def get_user_output_dir(user_name: str) -> Path:
    """社員名ごとの保存先フォルダを取得（無ければ自動生成）"""
    user_dir = BASE_DOCUMENTS_DIR / user_name
    user_dir.mkdir(parents=True, exist_ok=True)
    return user_dir