# WBS/EVM Ecosystem: AI-Driven Framework for Professional Project Management

WBS の構築から EVM 分析までを Python スクリプトで制御し、生成 AI による洞察を統合することで、プロジェクトマネジメントにおける「準備と検証」の工数を最小化する統合フレームワークです。

## 解決する課題 (Problem Statements)

本エコシステムは、エクセルベースの従来型 WBS/EVM 運用において、専門家が直面する以下の実務的課題を解決します。

*   **分析シートの作成コスト**: EVM グラフや複雑な計算式を内包した分析用シートの構築にかかる、膨大な手作業の排除。
*   **不透明なリスクの言語化**: 数値（CPI/SPI）の推移に潜む遅延予兆の特定と、PMBOK 理論に基づく具体的な示唆の生成。
*   **データの信頼性担保**: 行追加や編集に伴う参照エラー、計算不備の自動検証による「整合性の絶対遵守」。

## 主要機能 (Key Features)

### 1. Automated Analytics Engine (分析エンジンの自動化)
Python スクリプトにより、複雑な参照関係を持つ EVM 分析用エクセルシートを即時生成します。手作業によるセットアップをゼロにし、PM が分析そのものに集中できる環境を提供します。

### 2. AI Insight Generation (AI インサイト生成)
PMBOK 理論に基づき、AI が実績値の推移を解釈します。単なる数値報告に留まらず、「このままの推移では〇〇に影響する」といった具体的なリスク予兆を言語化し、レポートします。

### 3. Automated Integrity Guard (整合性ガード)
「Script-as-Master」思想に基づき、コマンド一つでエクセル内の計算不備や論理的矛盾を自動検証します。常に「壊れていないデータ」であることを保証し、意思決定の質を高めます。

## 設計思想 (Core Philosophy)

*   **Script-as-Master**: エクセルを「UI（出力）」と捉え、Python スクリプトを「真実のソース (SSOT)」として扱います。これにより、高度な再現性と自動修復を実現します。
*   **Theory-First**: すべての分析とアドバイスは、PMBOK 等の標準的なプロジェクトマネジメント理論を基盤としています。

## プロジェクト構成

*   `/scripts`: 分析、検証、生成のコアロジック。
*   `/templates`: WBS 構造およびスタイルの定義。
*   `/projects`: 各プロジェクトの実データ管理。

## インストール (Installation)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## クイックスタート

1.  **「型」を抽出する**:
    `./scripts/check_wbs.sh 理想のエクセル.xlsx`
    （`wbs_structure.json` が生成されます）

2.  **テンプレートを復元する**:
    `python scripts/generate_master_template.py --config wbs_structure.json --output 新規WBS.xlsx`
    （正しい数式が埋め込まれた新品のエクセルが生成されます）

3.  **整合性を確認する**:
    `python scripts/check_wbs_integrity.py 新規WBS.xlsx --sheet "WBS_EVM"`
