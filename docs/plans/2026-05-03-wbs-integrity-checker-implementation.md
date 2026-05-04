# WBS整合性チェックツール 実装計画

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、移植された `subagent-driven-development` スキル（推奨）または `executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。

**目標:** Excel WBSの論理的な矛盾（過負荷、日付逆転、実績不備）を検出し、ターミナルとExcelファイルに結果を出力するツールを構築する。

**アーキテクチャ:**
1. `pandas` と `openpyxl` を使用してExcelを読み書きする。
2. `jpholiday` を使用して日本の祝日を考慮した稼働日計算を行う。
3. チェックロジックを独立した関数として実装し、各行に対して適用する。
4. 結果をターミナルに表示し、Excelの「WBS_EVM」シートの右端に「整合性チェック結果」列を追加して保存する。

**技術スタック:**
- Python 3.x
- pandas, openpyxl, jpholiday, pytest

---

### タスク 1: 依存ライブラリの導入と基礎関数の実装

**ファイル:**
- 変更: なし (pip install)
- 作成: `scripts/utils.py` (祝日・稼働日計算用共通ユーティリティ)
- テスト: `tests/test_utils.py`

- [ ] **ステップ 1: 依存ライブラリのインストール**
    - 実行: `source .venv/bin/activate && pip install jpholiday`
- [ ] **ステップ 2: 稼働日計算ユーティリティの実装**
    - 土日祝日を除外した日数を返す関数 `get_working_days(start_date, end_date)` を作成。
- [ ] **ステップ 3: 失敗するテストの作成**
    - 祝日（例：2026/05/03-05/06のゴールデンウィーク）を含む期間の計算テストを書く。
- [ ] **ステップ 4: テストの実行とパス確認**
    - 実行: `pytest tests/test_utils.py`
- [ ] **ステップ 5: コミット**
    - `git add scripts/utils.py tests/test_utils.py && git commit -m "feat: add working day calculation utility with holiday support"`

### タスク 2: チェックスクリプトの骨格作成とExcel読み込み

**ファイル:**
- 作成: `scripts/check_wbs_integrity.py`
- テスト: `tests/test_check_wbs_integrity.py`

- [ ] **ステップ 1: スクリプトのベース実装**
    - 引数でExcelファイルのパスを受け取り、pandasで読み込む処理を実装。
- [ ] **ステップ 2: 基本的な検証関数の実装 (日付逆転チェック)**
    - `開始日 > 終了日` を検知するロジックを実装。
- [ ] **ステップ 3: テスト用ダミーExcelでの検証**
    - 不整合データを含む仮のExcel/DataFrameでロジックが動作することを確認。
- [ ] **ステップ 4: コミット**
    - `git add scripts/check_wbs_integrity.py && git commit -m "feat: skeleton of WBS integrity checker with date validation"`

### タスク 3: 高度な整合性チェックロジックの実装

**ファイル:**
- 変更: `scripts/check_wbs_integrity.py`

- [ ] **ステップ 1: 過負荷チェックの実装**
    - `稼働日数 < 予定工数` を検知するロジックを追加。
- [ ] **ステップ 2: フェーズ間順序性チェックの実装**
    - 作成終了 -> レビュー開始 -> 修正開始 の順序矛盾を検知。
- [ ] **ステップ 3: 実績整合性チェックの実装**
    - 進捗率100%時の終了日未入力、実績工数と日付の矛盾などを検知。
- [ ] **ステップ 4: 統合テストの実施**
    - 全てのチェック項目を網羅したテストを実行。
- [ ] **ステップ 5: コミット**
    - `git commit -am "feat: implement advanced validation logic for workload and phase sequence"`

### タスク 4: Excelへの書き戻し機能と最終統合

**ファイル:**
- 変更: `scripts/check_wbs_integrity.py`

- [ ] **ステップ 1: Excel出力機能の実装**
    - `openpyxl` を使用して、既存のレイアウトを維持したまま右端の列にエラーメッセージを書き込む。
- [ ] **ステップ 2: ターミナル出力の整形**
    - ユーザーに見やすい色付き、または表形式のサマリーを表示。
- [ ] **ステップ 3: 実ファイルでの最終動作確認**
    - `projects/test_project/hoge_wbs_evm.xlsx` に対して実行し、結果を確認。
- [ ] **ステップ 4: ドキュメント更新**
    - 使い方を README.md または新しいドキュメントに記載。
- [ ] **ステップ 5: 完了報告とコミット**
    - `git commit -am "feat: add excel write-back and final terminal output formatting"`
