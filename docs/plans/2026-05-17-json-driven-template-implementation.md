# JSON駆動型テンプレート復元システム 実装計画 (TDD & Review Edition)

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、移植された `subagent-driven-development` スキル（推奨）または `executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。
> **重要:** 各タスクの完了後、必ず `requesting-code-review` スキルを起動し、ユーザーまたはシニアエージェントのレビューを受けてください。

**目標:** 既存エクセルから抽出された JSON 設定ファイルを元に、任意のフェーズ構成と正しい数式を持つ WBS テンプレートを動的に生成する。

---

### タスク 1: WBSConfigManager のエクスポート機能の TDD 強化

**ファイル:**
- 変更: `scripts/wbs_config_manager.py`
- テスト: `tests/test_wbs_config_manager.py`

- [ ] **ステップ 1: [RED] 設定ファイルへの書き出しと再読み込みの整合性を検証するテストを追加**
```python
def test_save_and_load_config_integrity():
    # 1. infer_structure で設定を生成
    # 2. save_config で一時ファイルに保存
    # 3. 再度読み込んで、元の設定と一致するか、あるいは生存確認をパスするか検証
```
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `save_config` および `load_or_infer` の微修正を行い、テストをパスさせる**
- [ ] **ステップ 4: [REFACTOR] JSON 構造を整理し、生成器が参照しやすいようキー名を統一**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行**
- [ ] **ステップ 6: コミット**

---

### タスク 2: TemplateGenerator のリファクタリング（動的フェーズ対応）

**ファイル:**
- 変更: `scripts/generate_master_template.py`
- 新規テスト: `tests/test_template_generator_dynamic.py`

- [ ] **ステップ 1: [RED] 任意のフェーズ構成を含む JSON からエクセルを生成するテストを作成**
```python
def test_generate_from_custom_config():
    config = {
        "sheet_name": "Custom_WBS",
        "header_row": 2,
        "columns": {
            "common": {"id": {"index": 0, "name": "ID"}, "name": {"index": 1, "name": "Name"}},
            "phases": [
                {
                    "name": "Phase1",
                    "mapping": {"plan_start": {"index": 2, "name": "Start"}, "plan_effort": {"index": 3, "name": "Cost"}}
                }
            ]
        }
    }
    output = "tests/data/tmp_master.xlsx"
    gen = TemplateGenerator(output)
    gen.generate(config=config)
    assert os.path.exists(output)
```
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `TemplateGenerator` を引数対応にリファクタリング**
```python
class TemplateGenerator:
    def generate(self, config=None):
        if config:
            self._generate_from_config(config)
        else:
            self._generate_default()

    def _generate_from_config(self, config):
        # 設定に従ってシート作成とヘッダー配置までを実装
        pass
```
- [ ] **ステップ 4: [REFACTOR] 共通スタイル処理を動的列数に対応させる**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行**
- [ ] **ステップ 6: コミット**

---

### タスク 3: 役割ベースの数式動的生成ロジックの実装

**ファイル:**
- 変更: `scripts/generate_master_template.py`

- [ ] **ステップ 1: [RED] 動的に配置された列を正しく参照する数式が入っているか検証するテストを追加**
```python
def test_dynamic_formula_references():
    # 1. フェーズ内の EV 列、工数予定列、進捗率列のインデックスを取得
    # 2. 生成されたエクセルの EV セルの数式が = [工数列][行] * ([進捗率列][行]/100) となっているか検証
```
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] 数式動的組み立てロジックの実装**
```python
from openpyxl.utils import get_column_letter

def build_formula(mapping, role, row):
    if role == 'ev':
        cost_col = get_column_letter(mapping['plan_effort']['index'] + 1)
        prog_col = get_column_letter(mapping['progress']['index'] + 1)
        return f"={cost_col}{row} * ({prog_col}{row}/100)"
    # ... 他の役割も同様に実装
```
- [ ] **ステップ 4: [REFACTOR] 数式テンプレートを管理する専用メソッドに抽出**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行**
- [ ] **ステップ 6: コミット**

---

### タスク 4: CLI インターフェースの統合

**ファイル:**
- 変更: `scripts/generate_master_template.py`

- [ ] **ステップ 1: [RED] 引数 `--config` を受け取り生成を実行するテストを作成**
- [ ] **ステップ 2: テスト実行（失敗を確認）**
- [ ] **ステップ 3: [GREEN] `argparse` 実装**
- [ ] **ステップ 4: [REFACTOR] 既存のデフォルト挙動との互換性整理**
- [ ] **ステップ 5: [REVIEW] `requesting-code-review` を実行**
- [ ] **ステップ 6: コミット**

---

### タスク 5: 退行保護 (Regression Test) とクリーンアップ

- [ ] **ステップ 1: 既存の全テストを含む全テストの実行による副作用確認**
    - 実行コマンド: `pytest tests/ -v`
- [ ] **ステップ 2: テストデータのクリーンアップ**
- [ ] **ステップ 3: `docs/TODO.md` の更新**
- [ ] **ステップ 4: [FINAL REVIEW] `requesting-code-review` で全工程の完了を確認**
