"""Module for PowerPoint layout optimization (title alignment, foreground page numbering, and hanging indentation)."""

import logging
import pathlib
from pathlib import Path
import re
from typing import Union

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from summary_schema import WorkbookSummarySchema

logger = logging.getLogger(__name__)

FONT_JAPANESE = "メイリオ"
COLOR_PRIMARY = RGBColor(0, 51, 102)        # Dark Navy
COLOR_SECONDARY = RGBColor(235, 243, 250)   # Light Blue
COLOR_CARD_BG = RGBColor(248, 249, 250)     # Card Background
COLOR_CARD_BORDER = RGBColor(218, 224, 233)
COLOR_TEXT_MAIN = RGBColor(51, 51, 51)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_BADGE_BG = RGBColor(254, 249, 231)
COLOR_BADGE_BORDER = RGBColor(241, 196, 15)

SECTION_CONFIG = {
    "成果（事実）": {
        "title": "🎯 成果（事実 / 定量データ）",
        "color": RGBColor(0, 51, 102),
    },
    "評価（価値）": {
        "title": "💎 評価（ビジネス価値）",
        "color": RGBColor(39, 174, 96),
    },
    "課題（ギャップ）": {
        "title": "⚠️ 課題（ボトルネック）",
        "color": RGBColor(211, 84, 0),
    },
    "次の一手（アクション）": {
        "title": "🚀 次の一手（具体的対策）",
        "color": RGBColor(41, 128, 185),
    },
}

STATUS_COLORS = {
    "【対応中】": RGBColor(41, 128, 185),
    "【要検討】": RGBColor(211, 84, 0),
    "【未着手】": RGBColor(192, 57, 43),
    "【完了】": RGBColor(39, 174, 96),
    "【計画中】": RGBColor(142, 68, 173),
}


class PPTXGenerator:
    """Class responsible for generating optimized PowerPoint reports from structured summaries."""

    def __init__(self, template_path: Union[str, Path, None] = None):
        if template_path is None:
            from config import TEMPLATE_PATH
            template_path = TEMPLATE_PATH

        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(
                f"Template not found: {self.template_path.absolute()}"
            )

    def _clean_learning_info(self, raw_info: str) -> str:
        """Clean up raw learning or appeal information strings."""
        if not raw_info:
            return ""
        cleaned = re.sub(
            r"^(💡|自己学習・活用ツール|アピール|使用ツール・教材|教材|使用ツール|参照テキスト/サイト|参照テキスト)[\s:：・]*",
            "",
            raw_info,
        ).strip()
        cleaned = re.sub(
            r"(使用ツール・教材|教材|使用ツール)[\s:：]*", "", cleaned
        ).strip()
        return cleaned

    def _clean_bullet_text(self, text: str) -> str:
        """Clean up bullet point texts by stripping prefix tags and bullet symbols."""
        if not text:
            return ""
        cleaned = re.sub(
            r"^(成果|事実|評価|価値|課題|ギャップ|次の一手|アクション)[\s:：・]*",
            "",
            text,
        ).strip()
        cleaned = re.sub(r"^(・|\s|[-*])*", "", cleaned).strip()
        return cleaned

    def _add_page_number(self, slide, page_num: int):
        """Draw the page number as a text box in the bottom right corner of the slide."""
        tx_box = slide.shapes.add_textbox(
            Inches(11.0), Inches(7.0), Inches(1.5), Inches(0.3)
        )
        tf = tx_box.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = str(page_num)
        p.font.name = FONT_JAPANESE
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_MAIN
        p.alignment = PP_ALIGN.RIGHT

    def _add_styled_bullet_paragraph(
        self,
        tf,
        text: str,
        font_size: float = 9.0,
        is_bold: bool = False,
    ):
        """Add a formatted bullet point with hanging indentation and status tag coloring."""
        clean_text = self._clean_bullet_text(text)
        if not clean_text:
            return

        p = (
            tf.add_paragraph()
            if len(tf.paragraphs[0].text) > 0
            else tf.paragraphs[0]
        )
        p.line_spacing = 1.15
        p.space_after = Pt(3)

        p.margin_left = Pt(14)
        p.first_line_indent = Pt(-14)

        tag_match = re.search(r"【(対応中|要検討|未着手|完了|計画中)】", clean_text)
        if tag_match and tag_match.group(0) in STATUS_COLORS:
            tag_str = tag_match.group(0)
            rest_str = clean_text.replace(tag_str, "").strip()

            run_bullet = p.add_run()
            run_bullet.text = "・ "
            run_bullet.font.name = FONT_JAPANESE
            run_bullet.font.size = Pt(font_size)
            run_bullet.font.color.rgb = COLOR_TEXT_MAIN

            run_tag = p.add_run()
            run_tag.text = f"{tag_str} "
            run_tag.font.name = FONT_JAPANESE
            run_tag.font.size = Pt(font_size)
            run_tag.font.bold = True
            run_tag.font.color.rgb = STATUS_COLORS.get(tag_str, COLOR_PRIMARY)

            run_rest = p.add_run()
            run_rest.text = rest_str
            run_rest.font.name = FONT_JAPANESE
            run_rest.font.size = Pt(font_size)
            run_rest.font.bold = is_bold
            run_rest.font.color.rgb = COLOR_TEXT_MAIN
        else:
            run_bullet = p.add_run()
            run_bullet.text = "・ "
            run_bullet.font.name = FONT_JAPANESE
            run_bullet.font.size = Pt(font_size)
            run_bullet.font.color.rgb = COLOR_TEXT_MAIN

            run_text = p.add_run()
            run_text.text = clean_text
            run_text.font.name = FONT_JAPANESE
            run_text.font.size = Pt(font_size)
            run_text.font.bold = is_bold
            run_text.font.color.rgb = COLOR_TEXT_MAIN

    def generate(
        self,
        summary_data: WorkbookSummarySchema,
        output_path: Union[str, Path],
        author_name: str = "氏名",
    ) -> Path:
        """Generate a complete PowerPoint presentation based on the structured summary data."""
        prs = Presentation(str(self.template_path))

        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]

        content_layout = (
            prs.slide_layouts[3]
            if len(prs.slide_layouts) > 3
            else prs.slide_layouts[0]
        )

        # -----------------------------------------------------------------
        # 1. Title Slide (Unified Department and Author Placement)
        # -----------------------------------------------------------------
        title_slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Title (idx: 0)
        if 0 in [ph.placeholder_format.idx for ph in title_slide.placeholders]:
            title_shape = title_slide.placeholders[0]
            title_shape.text = re.sub(
                r"[\(（].*?[\)）]", "", summary_data.document_title
            ).strip()
            for p in title_shape.text_frame.paragraphs:
                p.font.name = FONT_JAPANESE
                p.font.size = Pt(28)
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY

        # Department & Author Name (placed inside idx: 15)
        if 15 in [
            ph.placeholder_format.idx for ph in title_slide.placeholders
        ]:
            dept_shape = title_slide.placeholders[15]
            dept_shape.text = (
                f"所属：エンジニアリング事業部\n氏名：{author_name or '氏名'}"
            )
            for p in dept_shape.text_frame.paragraphs:
                p.font.name = FONT_JAPANESE
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.color.rgb = COLOR_TEXT_MAIN
                p.alignment = PP_ALIGN.RIGHT

        # Clear redundant spare placeholder (idx: 14)
        if 14 in [
            ph.placeholder_format.idx for ph in title_slide.placeholders
        ]:
            title_slide.placeholders[14].text = ""

        # -----------------------------------------------------------------
        # 2. Summary Slide (Page 2)
        # -----------------------------------------------------------------
        summary_slide = prs.slides.add_slide(content_layout)
        self._add_page_number(summary_slide, 2)

        if 0 in [
            ph.placeholder_format.idx for ph in summary_slide.placeholders
        ]:
            t_shape = summary_slide.placeholders[0]
            t_shape.text = "エグゼクティブ・サマリー（主要成果とストーリー）"
            for p in t_shape.text_frame.paragraphs:
                p.font.name, p.font.size, p.font.bold = (
                    FONT_JAPANESE,
                    Pt(18),
                    True,
                )
                p.font.color.rgb = COLOR_PRIMARY

        top_card = summary_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            Inches(1.5),
            Inches(11.8),
            Inches(2.3),
        )
        top_card.fill.solid()
        top_card.fill.fore_color.rgb, top_card.line.color.rgb = (
            COLOR_CARD_BG,
            COLOR_CARD_BORDER,
        )
        tf_top = top_card.text_frame
        tf_top.word_wrap = True
        tf_top.margin_left = Inches(0.25)
        p_top_h = tf_top.paragraphs[0]
        p_top_h.text = "■ 定量成果・要約（3点）"
        p_top_h.font.name, p_top_h.font.bold, p_top_h.font.size, p_top_h.font.color.rgb = (
            FONT_JAPANESE,
            True,
            Pt(12),
            COLOR_PRIMARY,
        )
        p_top_h.space_after = Pt(6)

        for insight in getattr(summary_data, "key_insights", [])[:3]:
            self._add_styled_bullet_paragraph(
                tf_top, insight, font_size=9.5, is_bold=True
            )

        bottom_card = summary_slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            Inches(4.0),
            Inches(11.8),
            Inches(2.6),
        )
        bottom_card.fill.solid()
        bottom_card.fill.fore_color.rgb, bottom_card.line.color.rgb = (
            COLOR_SECONDARY,
            COLOR_CARD_BORDER,
        )
        tf_bot = bottom_card.text_frame
        tf_bot.word_wrap = True
        tf_bot.margin_left = Inches(0.25)
        p_bot_h = tf_bot.paragraphs[0]
        p_bot_h.text = "■ プロジェクト推進ストーリー"
        p_bot_h.font.name, p_bot_h.font.bold, p_bot_h.font.size, p_bot_h.font.color.rgb = (
            FONT_JAPANESE,
            True,
            Pt(12),
            COLOR_PRIMARY,
        )
        p_bot_h.space_after = Pt(6)

        for sent in re.split(
            r"(?<=[。！!？?])", getattr(summary_data, "overall_summary", "")
        ):
            if sent.strip():
                self._add_styled_bullet_paragraph(
                    tf_bot, sent.strip(), font_size=9.0
                )

        # -----------------------------------------------------------------
        # 3. Detail Slides (Page 3 and onwards)
        # -----------------------------------------------------------------
        current_page = 3
        for slide_data in summary_data.slides:
            slide = prs.slides.add_slide(content_layout)
            cat = getattr(slide_data, "category", "社内業務") or "社内業務"

            if 0 in [ph.placeholder_format.idx for ph in slide.placeholders]:
                t_shape = slide.placeholders[0]
                t_shape.text = f"【{cat}】 {slide_data.slide_title}"
                for p in t_shape.text_frame.paragraphs:
                    p.font.name, p.font.size, p.font.bold = (
                        FONT_JAPANESE,
                        Pt(18),
                        True,
                    )
                    p.font.color.rgb = COLOR_PRIMARY

            clean_info = self._clean_learning_info(
                getattr(slide_data, "learning_info", "")
            )
            grid_top_offset = Inches(1.6)
            card_height = Inches(2.4)

            if clean_info:
                label_prefix = (
                    "参照テキスト/サイト" if "自己学習" in cat else "使用ツール"
                )
                badge = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Inches(0.8),
                    Inches(1.4),
                    Inches(11.8),
                    Inches(0.45),
                )
                badge.fill.solid()
                badge.fill.fore_color.rgb, badge.line.color.rgb = (
                    COLOR_BADGE_BG,
                    COLOR_BADGE_BORDER,
                )
                tf_b = badge.text_frame
                tf_b.word_wrap = True
                p_b = tf_b.paragraphs[0]
                p_b.text = (
                    f"💡 アピールポイント ｜ {label_prefix}：{clean_info}"
                )
                p_b.font.name, p_b.font.bold, p_b.font.size = (
                    FONT_JAPANESE,
                    True,
                    Pt(10.0),
                )
                p_b.font.color.rgb = COLOR_PRIMARY
                grid_top_offset, card_height = Inches(1.95), Inches(2.25)

            four_sections = {
                "成果（事実）": [],
                "評価（価値）": [],
                "課題（ギャップ）": [],
                "次の一手（アクション）": [],
            }

            for item in getattr(slide_data, "bullet_points", []):
                cleaned_item = self._clean_bullet_text(item)
                if not cleaned_item:
                    continue

                if any(k in item for k in ["評価", "価値"]):
                    four_sections["評価（価値）"].append(cleaned_item)
                elif any(k in item for k in ["課題", "ギャップ"]):
                    four_sections["課題（ギャップ）"].append(cleaned_item)
                elif any(k in item for k in ["次の一手", "今後", "計画"]):
                    four_sections["次の一手（アクション）"].append(cleaned_item)
                else:
                    four_sections["成果（事実）"].append(cleaned_item)

            for act in getattr(slide_data, "action_items", []):
                cleaned_act = self._clean_bullet_text(act)
                if (
                    cleaned_act
                    and cleaned_act
                    not in four_sections["次の一手（アクション）"]
                ):
                    four_sections["次の一手（アクション）"].append(cleaned_act)

            grid_params = [
                ("成果（事実）", Inches(0.8), grid_top_offset),
                ("評価（価値）", Inches(6.8), grid_top_offset),
                (
                    "課題（ギャップ）",
                    Inches(0.8),
                    grid_top_offset + card_height + Inches(0.15),
                ),
                (
                    "次の一手（アクション）",
                    Inches(6.8),
                    grid_top_offset + card_height + Inches(0.15),
                ),
            ]

            for sec_name, gx, gy in grid_params:
                cfg = SECTION_CONFIG[sec_name]
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, gx, gy, Inches(5.8), card_height
                )
                card.fill.solid()
                card.fill.fore_color.rgb, card.line.color.rgb = (
                    COLOR_CARD_BG,
                    cfg["color"],
                )
                card.line.width = Pt(1.5)

                header = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    gx,
                    gy,
                    Inches(5.8),
                    Inches(0.38),
                )
                header.fill.solid()
                header.fill.fore_color.rgb = cfg["color"]
                header.line.fill.background()

                tf_h = header.text_frame
                p_h = tf_h.paragraphs[0]
                p_h.text = cfg["title"]
                p_h.font.name, p_h.font.bold, p_h.font.size = (
                    FONT_JAPANESE,
                    True,
                    Pt(10.0),
                )
                p_h.font.color.rgb = COLOR_WHITE
                p_h.alignment = PP_ALIGN.CENTER

                tf_c = card.text_frame
                tf_c.word_wrap = True
                tf_c.margin_left, tf_c.margin_right, tf_c.margin_top = (
                    Inches(0.2),
                    Inches(0.2),
                    Inches(0.45),
                )

                items_to_draw = four_sections[sec_name]
                if items_to_draw:
                    for item in items_to_draw:
                        self._add_styled_bullet_paragraph(
                            tf_c, item, font_size=8.5
                        )

            # Draw page number on top of each slide
            self._add_page_number(slide, current_page)
            current_page += 1

        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out_path))
        return out_path