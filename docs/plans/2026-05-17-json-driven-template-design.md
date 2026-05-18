# Design Document: JSON駆動型テンプレート復元システム

## 1. 目的
「人間が使いやすく調整したエクセル」から構造（型）を抽出し、それを設計図（JSON）として新しいプロジェクト用の高品質なテンプレートを再生成する。

## 2. 背景と課題
- **現状**: `generate_master_template.py` が 3 フェーズ固定の構造をハードコードしており、プロジェクトごとの微調整を反映できない。
- **理想**: 特定のプロジェクトで「育った」エクセル構造を簡単に再利用可能にし、かつ数式やスタイルの整合性をシステム側で保証する。

## 3. システムアーキテクチャ
### 3.1. コンポーネント構成
1. **Extractor (既存拡張)**: `WBSConfigManager` を使用し、既存エクセルから `wbs_structure.json` を出力。
2. **Generator (新規/リファクタ)**: `generate_master_template.py` が JSON を読み込み、エクセルを構築。

### 3.2. データフロー
```mermaid
graph LR
    A[既存Excel] -->|Scan| B(WBSConfigManager)
    B -->|Export| C[wbs_structure.json]
    C -->|Input| D(TemplateGenerator)
    D -->|Build| E[新MasterTemplate.xlsx]
```

## 4. 詳細仕様
### 4.1. 設定ファイル (wbs_structure.json) の拡張
従来の構成に加え、復元に必要な情報を網羅する。
- `sheet_name`: ターゲットシート名。
- `columns.common`: ID, 名称などの全フェーズ共通項目。
- `columns.phases`: 各フェーズの名前と、その中での各役割（plan_start等）の列位置。

### 4.2. 生成エンジンの動的化
- **フェーズ・カラムの動的配置**: `columns.phases` の要素数分だけフェーズ見出しを作成し、各役割を `index` に基づいて配置する。
- **数式の動的解決**:
    - JSON 内の役割（例: `plan_effort` と `progress`）のインデックスを特定。
    - 列番号をアルファベット（A, B, C...）に変換し、相対参照の数式を組み立てる。
    - 例: `EV = COLUMN(plan_effort) * (COLUMN(progress)/100)`

### 4.3. CLI インターフェース
`scripts/generate_master_template.py` に以下の引数を追加する。
- `--config <path>`: 使用する設定ファイル（JSON）のパス。
- `--output <path>`: 生成するエクセルの保存先。
- `--sheet-name <name>`: (任意) シート名の指定。省略時は JSON の値を採用。

## 5. 成功基準 (DoD)
- [ ] 任意のフェーズ数（1〜5程度）を含む JSON からエクセルが正しく生成されること。
- [ ] 生成されたエクセルの PV/EV/AC 計算式が、動的に配置された列を正しく参照していること。
- [ ] スタイル（罫線、ヘッダー色）が、従来のテンプレートと同等の品質で適用されていること。

## 6. リスクと対策
- **数式エラー**: 列数が大幅に変わった際に SUMIFS 等の集計範囲がズレる可能性がある。
  - 対策: `WBSConfigManager` の推論ロジックと共通の役割定義を使用し、テストコードで数式の参照先を検証する。
