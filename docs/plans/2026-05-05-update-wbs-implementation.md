# 対話型WBS更新機能 (update-wbs) 実装計画

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、`executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。また、各タスクで指定された補助スキルを必ず起動してください。

**目標:** 自然言語および `/update-wbs` コマンドで動作し、ステートフルなフェーズ判別と物理的整合性チェックを行う専用スキルを構築する。

**アーキテクチャ:** 
1. Gemini CLI のカスタムスキルとして定義。
2. バックエンドに `openpyxl` を使用した Python スクリプトを配置。
3. 既存の `scripts/check_wbs_integrity.py` を再利用して保存直前に検証を行う。

**技術スタック:** Python 3.12, pandas, openpyxl, pytest, Gemini CLI Skill system

---

### タスク 1: スキル定義の作成 (規律の構築)

**使用スキル:** `writing-skills`
**ファイル:**
- 作成: `.gemini/skills/update-wbs/SKILL.md`

- [ ] **ステップ 1: スキルディレクトリの作成**
  実行: `mkdir -p .gemini/skills/update-wbs`
- [ ] **ステップ 2: SKILL.md の記述**
  - 自然言語での発動条件（進捗報告など）を `description` に記述。
  - 承認フロー（はい/いいえ）の義務化を規律として記述。
  - バックエンドスクリプト `scripts/update_wbs.py` への橋渡しを記述。
- [ ] **ステップ 3: コミット**
  実行: `git add .gemini/skills/update-wbs/SKILL.md && git commit -m "feat: update-wbs スキル定義の作成"`

---

### タスク 2: バックエンドロジックの TDD 実装 (RED: 失敗テスト)

**使用スキル:** `test-driven-development`
**ファイル:**
- 作成: `tests/test_update_wbs.py`
- 作成: `scripts/update_wbs.py` (空の関数のみ)

- [ ] **ステップ 1: RED - 失敗するテストを作成**
  - 内容: 特定の機能ID、フェーズ、工数を指定して Excel が更新されることを期待するテスト。
- [ ] **ステップ 2: テストの実行と失敗の確認**
  実行: `.venv/bin/pytest tests/test_update_wbs.py`
  期待値: FAIL (ImportError or NotImplementedError)

---

### タスク 3: バックエンドロジックの TDD 実装 (GREEN: 最小実装)

**使用スキル:** `test-driven-development`
**ファイル:**
- 変更: `scripts/update_wbs.py`

- [ ] **ステップ 1: GREEN - 最小限の実装を作成**
  - ステートフルなフェーズ特定ロジックの実装。
  - `openpyxl` によるセル書き込み。
  - `check_wbs_integrity` の呼び出しと例外処理。
- [ ] **ステップ 2: テストの実行とパスの確認**
  実行: `.venv/bin/pytest tests/test_update_wbs.py`
  期待値: PASS
- [ ] **ステップ 3: リファクタリング**
  - 重複コードの整理、コメントの追加。
- [ ] **ステップ 4: コミット**
  実行: `git add scripts/update_wbs.py tests/test_update_wbs.py && git commit -m "feat: update-wbs バックエンドロジックの実装"`

---

### タスク 4: コードレビューと最終検証

**使用スキル:** `requesting-code-review`, `verification-before-completion`

- [ ] **ステップ 1: コードレビューの依頼**
  - `code-reviewer` サブエージェントを派遣し、設計ドキュメントとの整合性を確認。
- [ ] **ステップ 2: 実機検証 (End-to-End)**
  - 実際に `/update-wbs` 相当の操作を CLI 経由でシミュレートし、`projects/test_project/hoge_wbs_evm.xlsx` が正しく更新されるか確認。
- [ ] **ステップ 3: 完了の宣言**
  - `verification-before-completion` に基づき、全てのテストパスと整合性チェックのパスを証拠として提示。
