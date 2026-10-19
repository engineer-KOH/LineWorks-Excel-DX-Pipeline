# 🚀 LineWorks-Excel-DX-Pipeline

## 📌 概要（Overview）
このシステムは、LINE WORKSのWebhook連携およびExcel処理を基軸に、日々提出される日報データの収集・要約・ドキュメント化を効率化・自動化するためのPythonツールです。  
手動で行っていた煩雑なデータ処理をパイプライン化し、業務効率化と正確性の向上を実現します。

## 🎯 主な目的（Objective）
- **業務効率化**: 日報の転記や集計、要約にかかる工数を大幅に削減
- **ヒューマンエラー防止**: 手入力や転記ミス、データ抜けの排除
- **AI・外部連携**: LINE WORKS WebhookとGemini APIを活用したスマートな自動処理

## 🛠️ 使用技術（Tech Stack）
- **言語**: Python
- **主要ライブラリ**: `openpyxl`, `python-pptx`, `google-genai`, `pandas`, `fastapi` など
- **インフラ/外部サービス**: Gemini API, LINE WORKS Webhook

## ⚙️ システム構成・処理フロー（Workflow）
- **入力 (Inputs)**: LINE WORKSからのWebhook受信、または指定フォルダ内のExcelデータ読み込み
- **処理 (Process)**: データの正規化、バリデーション、AI（Gemini API）によるテキスト要約・構造化
- **出力 (Outputs)**: 整理されたレポートや成果物ドキュメントの自動生成

## 🚀 使い方（Quick Start）
- **1.リポジトリのクローン**: 「git clone [https://github.com/mochapaparocku-crypto/LineWorks-Excel-DX-Pipeline.git](https://github.com/mochapaparocku-crypto/LineWorks-Excel-DX-Pipeline.git)」
- **2.依存ライブラリのインストール**: 「pip install -r requirements.txt」
- **3.スクリプトの実行**: 「python コード一式/main.py」

## 📁 フォルダ構成（Repository Structure）
```text
LineWorks-Excel-DX-Pipeline/
├── コード一式/           # メインのPythonスクリプトおよび設定ファイル群
│   ├── main.py          # エントリーポイント
│   ├── config.py        # 設定管理
│   └── ...              # 各種モジュール
└── .gitignore           # 機密情報・生成物の除外設定
