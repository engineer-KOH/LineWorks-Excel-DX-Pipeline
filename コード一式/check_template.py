"""
==============================================================================
モジュール名 : check_template.py
機能概要     : aspark_template.pptx のレイアウト構造およびプレースホルダー確認
==============================================================================
"""
from pptx import Presentation
from pathlib import Path

def inspect_pptx_template(template_path: str):
    path = Path(template_path)
    if not path.exists():
        print(f"エラー: テンプレートファイルが見つかりません: {path.absolute()}")
        return

    prs = Presentation(template_path)
    print(f"=== テンプレート確認: {path.name} ===")
    print(f"スライド幅: {prs.slide_width.cm:.2f} cm")
    print(f"スライド高さ: {prs.slide_height.cm:.2f} cm")
    print(f"定義済みレイアウト数: {len(prs.slide_layouts)}\n")

    for idx, layout in enumerate(prs.slide_layouts):
        print(f"--- Layout [{idx}]: {layout.name} ---")
        for shape in layout.shapes:
            if shape.is_placeholder:
                ph = shape.placeholder_format
                print(f"  Placeholder Index: {ph.idx} | Type: {ph.type} | Name: {shape.name}")

if __name__ == "__main__":
    inspect_pptx_template("aspark_template.pptx")