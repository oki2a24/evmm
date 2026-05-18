# JSON駆動型テンプレート復元システム 実装計画 (TDD & Review Edition)

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、移植された `subagent-driven-development` スキル（推奨）または `executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。
> **重要:** 各タスクの完了後、必ず `requesting-code-review` スキルを起動し、ユーザーまたはシニアエージェントのレビューを受けてください。

**目標:** 既存エクセルから抽出された JSON 設定ファイルを元に、任意のフェーズ構成と正しい数式を持つ WBS テンプレートを動的に生成する。

**アーキテクチャ:**
1. **Extractor**: `WBSConfigManager` を拡張し、スキャンした構造を `wbs_structure.json` としてエクスポートする機能を強化。
2. **Generator**: `TemplateGenerator` クラスをリファクタリングし、役割（Role）ベースで列位置を決定し、Excel 数式を動的に組み立てるロジックを導入。

**技術スタック:** Python, openpyxl, pandas, pytest

---

### タスク 1: WBSConfigManager のエクスポート機能の TDD 強化

**ファイル:**
- 変更: `scripts/wbs_config_manager.py`
- テスト: `tests/test_wbs_config_manager.py`

- [ ] **ステップ 1: [RED] 設定ファイルへの書き出しと再読み込みの整合性を検証するテストを追加**
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `save_config` および `load_or_infer` の微修正を行い、テストをパスさせる**
- [ ] **ステップ 4: [REFACTOR] JSON 構造の可読性や例外処理を整理**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行し、マッピングの完全性を確認**
- [ ] **ステップ 6: コミット**

---

### タスク 2: TemplateGenerator のリファクタリング（動的フェーズ対応）

**ファイル:**
- 変更: `scripts/generate_master_template.py`
- 新規テスト: `tests/test_template_generator_dynamic.py`

- [ ] **ステップ 1: [RED] 任意のフェーズ構成（例: 2フェーズ）を含む JSON からエクセルを生成し、シート名とヘッダーを検証するテストを作成**
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `TemplateGenerator` をリファクタリングし、JSON 設定に基づいた基本的なシート構築を実装**
- [ ] **ステップ 4: [REFACTOR] 罫線適用（_apply_borders）などの共通スタイル処理を動的列数に対応させる**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行し、動的配置ロジックの汎用性を確認**
- [ ] **ステップ 6: コミット**

---

### タスク 3: 役割ベースの数式動的生成ロジックの実装

**ファイル:**
- 変更: `scripts/generate_master_template.py`

- [ ] **ステップ 1: [RED] 生成されたエクセルの PV/EV/AC セルに、正しい列をターゲットとした数式が入っているか検証するテストを追加**
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `openpyxl.utils.get_column_letter` を活用し、役割（Role）名をキーに数式文字列を動的に組み立てるロジックを実装**
- [ ] **ステップ 4: [REFACTOR] 数式テンプレート（PV計算式等）の保守性を高めるため、定数化またはメソッド化を検討**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行し、数式の相対参照の正確性を確認**
- [ ] **ステップ 6: コミット**

---

### タスク 4: CLI インターフェースと既存統合の検証

**ファイル:**
- 変更: `scripts/generate_master_template.py`

- [ ] **ステップ 1: [RED] 引数 `--config` を受け取り、正しく生成が完了することを検証する統合テスト（またはサブプロセス実行テスト）を作成**
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `argparse` による引数処理を実装し、テストをパスさせる**
- [ ] **ステップ 4: [REFACTOR] 既存のデフォルト挙動（引数なし実行）との互換性を確保しつつコードを整理**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行し、CLIのユーザビリティを確認**
- [ ] **ステップ 6: コミット**

---

### タスク 5: 退行保護 (Regression Test) とクリーンアップ

- [ ] **ステップ 1: 既存の全テスト (`tests/`) を実行し、今回の変更による既存機能への副作用がないことを証明する**
    - 実行コマンド: `pytest tests/ -v`
- [ ] **ステップ 2: 生成された一時的なエクセルファイルやテストデータのクリーンアップ**
- [ ] **ステップ 3: `docs/TODO.md` を更新し、次回の「学習機能」への布石を打つ**
- [ ] **ステップ 4: [FINAL REVIEW] 全工程の完了を `requesting-code-review` で報告し、セッションを締めくくる**
