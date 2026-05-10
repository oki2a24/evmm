# 自己進化型メンタリング・システム 実装計画 (Quality Guaranteed Edition)

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには `executing-plans` スキルを `activate_skill` で起動して使用してください。
> **厳格な規律:** 各タスクにおいて、実装前に必ず失敗するテストを書き (TDD)、実装後にテストがパスすることを確認し、最後に自己レビュー（コード・ドキュメントの整合性チェック）を物理的証拠と共に行うこと。

**目標:** PJ背景知識 (`context.md`) を活用した「伴走型メンター」の実装。新PJ開始時の自動生成機能を備え、高品質なコードと一貫性のあるドキュメントを維持する。

---

### タスク 1: コンテキスト基盤の自動生成ロジックの実装 (TDD & Review)

**ファイル:**
- 変更: `scripts/utils.py`
- テスト: `tests/test_utils.py`

- [ ] **ステップ 1: [TDD] 失敗するテストを作成**
`tests/test_utils.py` に `test_ensure_project_context_creates_dir_and_file` を追加。存在しないパスを渡し、`docs/context.md` が生成されない状態で実行し、失敗を確認する。

- [ ] **ステップ 2: [Act] 最小限の実装を作成**
`scripts/utils.py` に `ensure_project_context(project_path)` を実装。

- [ ] **ステップ 3: [Validate] テストをパスさせる**
`pytest tests/test_utils.py` を実行し、ファイルが期待通りの初期内容で生成されることを物理的に証明する。

- [ ] **ステップ 4: [Review] コード & 整合性チェック**
`utils.py` の他の関数とのスタイル不整合がないか、例外処理は適切かを確認し、結果を報告する。

- [ ] **ステップ 5: コミット**
`git add scripts/utils.py tests/test_utils.py && git commit -m "feat: プロジェクトコンテキスト基盤の自動生成ロジックを追加"`

---

### タスク 2: 分析プロセスへの自動初期化の統合 (Integration & Review)

**ファイル:**
- 変更: `scripts/analyze_evm.py`
- テスト: `tests/test_analyze_evm.py`

- [ ] **ステップ 1: [TDD] 統合テストの追加**
`analyze_project` を呼び出した際、`metadata` に `context_path` が含まれ、かつ実際にファイルが生成されていることを検証するテストを追加し、失敗を確認。

- [ ] **ステップ 2: [Act] analyze_project への統合**
`ensure_project_context` を呼び出し、結果をメタデータに格納するよう修正。

- [ ] **ステップ 3: [Validate] 全テストの実行**
`pytest tests/test_analyze_evm.py tests/test_utils.py` を実行し、統合に問題がないことを証明。

- [ ] **ステップ 4: [Review] ドキュメント整合性チェック**
`docs/plans/` 内の設計書と実装に乖離がないか再確認し、報告する。

- [ ] **ステップ 5: コミット**
`git add scripts/analyze_evm.py tests/test_analyze_evm.py && git commit -m "feat: 分析プロセスにコンテキストの自動初期化を統合"`

---

### タスク 3: AI アナリストへの運用規律の永続化 (Doc Integrity)

**ファイル:**
- 変更: `.gemini/observations/analyze-evm.md`
- 変更: `.gemini/observations/update-wbs.md`

- [ ] **ステップ 1: [Act] Observations の更新**
設計書に基づき、AI がコンテキストを読み書きする指示を具体的に記述する。

- [ ] **ステップ 2: [Validate] 整合性検証**
更新した Observations を読み込み、AI 自身が「新しい規律を正しく理解した」ことを、自分自身の言葉で要約して証明する。

- [ ] **ステップ 3: コミット**
`git add .gemini/observations/ && git commit -m "feat: AIアナリストにコンテキスト運用の規律を追加"`

---

### タスク 4: 動作実証 (End-to-End Validation)

- [ ] **ステップ 1: [Validate] 統合シナリオテスト**
1. `projects/e2e_test_pj/` ディレクトリを作成。
2. Excel をコピー。
3. `analyze-evm` スキル（またはスクリプト）を実行。
4. `docs/context.md` が生成され、AI がその内容を引用して問いかけを行うことを物理的に確認（ログを提示）。

- [ ] **ステップ 2: [Review] 最終品質チェック**
全ての要件が満たされ、負債（不要なプリント文やコメント）が残っていないか最終確認。

- [ ] **ステップ 3: コミット & 完了報告**
`git commit -m "chore: メンタリング・システムの構築完了と実証済みの報告"`
