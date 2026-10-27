"""Utility functions for filename sanitization, unique path generation, and author extraction."""

from datetime import datetime
from pathlib import Path
import re


def sanitize_filename(filename: str) -> str:
    """Replace invalid characters for Windows/Linux with safe alternatives."""
    invalid_chars = r'[\\/:*?"<>|]'
    return re.sub(invalid_chars, "_", filename)


def get_unique_filepath(
    output_dir: Path,
    base_name: str,
    extension: str = ".pptx",
    add_date: bool = True,
) -> Path:
    """Generate a sanitized safe filename and append a sequential number if it already exists."""
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
    """Extract author or personnel name from the given filename."""
    parts = Path(filename).stem.split("_")
    if len(parts) >= 2:
        return parts[1]

    match = re.search(r"[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]+", filename)
    if match:
        return match.group(0)

    return "担当者"


def get_user_output_dir(user_name: str, base_documents_dir: Path) -> Path:
    """Retrieve the output directory for each employee, creating it if it does not exist."""
    user_dir = base_documents_dir / user_name
    user_dir.mkdir(parents=True, exist_ok=True)
    return user_dir