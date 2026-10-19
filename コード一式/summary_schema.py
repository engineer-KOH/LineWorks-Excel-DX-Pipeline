"""
モジュール名 : summary_schema.py
機能概要     : Gemini APIのStructured Outputs向けPydanticスキーマ定義
"""
from typing import List, Optional
from pydantic import BaseModel, Field


class SlideSummarySchema(BaseModel):
    """個別スライド・セクションごとの要約構造"""
    category: str = Field(..., description="業務カテゴリ（例: '自己学習', '社内業務', '社内学習' など）")
    slide_title: str = Field(..., description="スライドまたはセクションのタイトル")
    learning_info: Optional[str] = Field(default=None, description="自己学習（参照テキスト/章）または社内学習（使用ツール）のアピール情報")
    bullet_points: List[str] = Field(..., description="主要なポイントの箇条書き")
    keywords: List[str] = Field(..., description="抽出された重要キーワードのリスト")
    action_items: Optional[List[str]] = Field(default=None, description="スライド単位のネクストアクション・対応事項")


class WorkbookSummarySchema(BaseModel):
    """ワークブック全体および個別シートの統合要約スキーマ"""
    document_title: str = Field(..., description="ドキュメント全体の表題や目的")
    overall_summary: str = Field(..., description="全体要約（200文字程度で簡潔に記述）")
    slides: List[SlideSummarySchema] = Field(..., description="シート単位またはセクション単位のスライド要約リスト")
    key_insights: List[str] = Field(..., description="重要な洞察や学びの箇条書きリスト")
    next_actions: Optional[List[str]] = Field(default=None, description="今後のタスクや推奨アクションのリスト")