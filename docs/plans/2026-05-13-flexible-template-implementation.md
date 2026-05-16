# 柔軟なWBSマッピングシステム 実装計画

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、移植された `subagent-driven-development` スキル（推奨）または `executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。

**目標:** エクセルの構造（列名、位置、フェーズ数）を動的に解決し、`wbs_structure.json` で管理する仕組みを導入することで、テンプレートの変更に強いWBS管理システムを実現する。

**アーキテクチャ:** 
1. `WBSConfigManager` クラスがエクセルのスキャンと推論、JSONの読み書きを担当する。
2. 既存のスクリプトは、直接列名を探すのではなく `WBSConfigManager` から「役割に応じた列インデックス」を取得する。
3. 構造が不明な場合や変更された場合は、対話的にユーザーに確認を求める。

**技術スタック:** Python (pandas, openpyxl), JSON

---

### ファイル構成
- `scripts/wbs_config_manager.py`: (新設) マッピング管理のコアロジック。
- `scripts/utils.py`: (変更) 共通ユーティリティの整理。
- `scripts/update_wbs.py`: (変更) マッピング対応へのリファクタリング。
- `scripts/analyze_evm.py`: (変更) マッピング対応へのリファクタリング。
- `scripts/check_wbs_integrity.py`: (変更) マッピング対応へのリファクタリング。
- `tests/test_wbs_config_manager.py`: (新設) マッピング管理のテスト。

---

### タスク 1: WBSConfigManager の基盤実装

**ファイル:**
- 作成: `scripts/wbs_config_manager.py`
- テスト: `tests/test_wbs_config_manager.py`

- [ ] **ステップ 1: 失敗するテストを作成**
  - エクセルファイルを読み込み、デフォルトの役割に対応するカラム名を推論できることを確認するテスト。
- [ ] **ステップ 2: テストが失敗することを確認するために実行**
- [ ] **ステップ 3: WBSConfigManager の最小限の実装**
  - `header_row` スキャン機能、正規表現による役割の推論ロジックを実装。
- [ ] **ステップ 4: テストがパスすることを確認するために実行**
- [ ] **ステップ 5: コミット**

### タスク 2: wbs_structure.json の永続化と対話型プロンプトの実装

**ファイル:**
- 変更: `scripts/wbs_config_manager.py`

- [ ] **ステップ 1: JSONの保存・読み込み機能を実装**
- [ ] **ステップ 2: ユーザーへの確認プロンプト（CLI）を実装**
  - `input()` を使用し、推論結果に対して Y/N または直接入力を受け取れるようにする。
- [ ] **ステップ 3: 構造変更の検知ロジックの実装**
  - 保存されたJSONのカラム名が現存するかチェックし、なければ再マッピングを促す。
- [ ] **ステップ 4: テストと手動検証**
- [ ] **ステップ 5: コミット**

### タスク 3: update_wbs.py のリファクタリング

**ファイル:**
- 変更: `scripts/update_wbs.py`

- [ ] **ステップ 1: update_wbs.py で WBSConfigManager を使用するように変更**
  - ハードコードされた `get_col_idx` 呼び出しを、`manager.get_column_index(role, phase)` に置き換える。
- [ ] **ステップ 2: 既存のテスト `tests/test_update_wbs.py` を実行してデグレードがないか確認**
- [ ] **ステップ 3: コミット**

### タスク 4: analyze_evm.py および check_wbs_integrity.py のリファクタリング

**ファイル:**
- 変更: `scripts/analyze_evm.py`
- 変更: `scripts/check_wbs_integrity.py`

- [ ] **ステップ 1: analyze_evm.py で WBSConfigManager を使用するように変更**
- [ ] **ステップ 2: check_wbs_integrity.py で WBSConfigManager を使用するように変更**
- [ ] **ステップ 3: 関連テスト `tests/test_analyze_evm.py`, `tests/test_check_wbs_integrity.py` を実行**
- [ ] **ステップ 4: コミット**

### タスク 5: 統合テストとクリーンアップ

- [ ] **ステップ 1: 全テストの一斉実行**
- [ ] **ステップ 2: 不要になった古い `get_col_idx` 等のコードを削除**
- [ ] **ステップ 3: README.md またはドキュメントの更新**
- [ ] **ステップ 4: 最終コミット**
