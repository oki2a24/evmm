# EVM 自動分析エンジン (EVMAnalyst) 実装計画

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: この計画をタスクごとに実装するには、`subagent-driven-development` スキル（推奨）または `executing-plans` スキルを `activate_skill` で起動して使用してください。ステップには追跡用のチェックボックス (`- [ ]`) を使用します。
> 各タスクでは、**TDD (Red-Green-Refactor)** を徹底し、**code-reviewer** による検証を経てから次へ進んでください。

**目標:** WBS Excelデータから CPI/SPI を算出し、将来予測を含む AI レポートを生成・Excel 反映する。

**アーキテクチャ:** 
- `scripts/analyze_evm.py` を核とし、`WBSIntegrityChecker` との連携による「規律ある分析」を実現する。
- 計算ロジックは Python で正確に行い、レポート生成に LLM のコンテキストを活用する。
- 既存の `scripts/utils.py` を再利用し、稼働日ベースの PV 計算を行う。

**技術スタック:** Python, pandas, openpyxl, pytest

---

### タスク 1: 分析エンジンの骨格と TDD 環境のセットアップ

**ファイル:**
- 作成: `scripts/analyze_evm.py`
- テスト: `tests/test_analyze_evm.py`

- [ ] **ステップ 1: 失敗するテストを作成 (RED)**
  - 指定された基準日に基づき、PV (線形按分) が正しく計算されることを検証するテスト。
- [ ] **ステップ 2: テストが失敗することを確認**
  - `export PYTHONPATH=$PYTHONPATH:. && .venv/bin/pytest tests/test_analyze_evm.py`
- [ ] **ステップ 3: 最小限の実装を作成 (GREEN)**
  - `EVMAnalyst` クラスの定義と、`calculate_pv` メソッドの実装。
- [ ] **ステップ 4: テストがパスすることを確認**
- [ ] **ステップ 5: code-reviewer 召喚**
  - ロジックの正確性と PMBOK 準拠性をレビュー。
- [ ] **ステップ 6: コミット**

### タスク 2: EV (出来高) および AC (実績コスト) の算出ロジック実装

**ファイル:**
- 変更: `scripts/analyze_evm.py`
- テスト: `tests/test_analyze_evm.py`

- [ ] **ステップ 1: 失敗するテストを作成 (RED)**
  - 進捗率から EV を、工数実績から AC を算出するロジックの検証。
- [ ] **ステップ 2: テストが失敗することを確認**
- [ ] **ステップ 3: 実装 (GREEN)**
  - `calculate_ev`, `calculate_ac` メソッドの実装。
- [ ] **ステップ 4: テストがパスすることを確認**
- [ ] **ステップ 5: code-reviewer 召喚**
- [ ] **ステップ 6: コミット**

### タスク 3: 整合性チェックとの統合 (AI Analyst Workflow)

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: `WBSIntegrityChecker` との連携テスト作成 (RED)**
  - 整合性エラーがある場合に警告または中断する挙動の検証。
- [ ] **ステップ 2: 連携ロジックの実装 (GREEN)**
  - 分析前に `WBSIntegrityChecker.run()` を実行するフローを追加。
- [ ] **ステップ 3: テストパス確認**
- [ ] **ステップ 4: code-reviewer 召喚**
- [ ] **ステップ 5: コミット**

### タスク 4: Excel への外科的更新ロジックの実装

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: Excel 更新テスト作成 (RED)**
  - `openpyxl` を使い、既存書式を維持したまま PV/EV/AC 列を更新できるか検証。
- [ ] **ステップ 2: `write_results_to_excel` の実装 (GREEN)**
  - `WBSIntegrityChecker` の同様のロジックを参考に、EVM カラムを更新。
- [ ] **ステップ 3: 実証テスト (TDD継続)**
  - テスト用 Excel を用いて、実際に値が正しく書き込まれることを確認。
- [ ] **ステップ 4: code-reviewer 召喚**
- [ ] **ステップ 5: コミット**

### タスク 5: AI レポート (予測分析) 生成機能の実装

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: レポート生成ロジックのテスト作成 (RED)**
  - CPI/SPI に基づく EAC (完成時予測) 計算の正確性を検証。
- [ ] **ステップ 2: `generate_ai_report` メソッドの実装 (GREEN)**
  - PMBOK 理論に基づいた診断メッセージの生成。
- [ ] **ステップ 3: テストパス確認**
- [ ] **ステップ 4: code-reviewer 召喚**
- [ ] **ステップ 5: コミット**

### タスク 6: 統合検証と実プロジェクトデータへの適用

- [ ] **ステップ 1: 最終統合テスト**
  - `projects/test_project/hoge_wbs_evm.xlsx` に対して 4/3 基準で分析を実行。
- [ ] **ステップ 2: 出力結果の目視確認と修正 (TDD)**
  - 修正が必要な場合は、まずテストコードを修正してから本体を直す。
- [ ] **ステップ 3: code-reviewer による最終承認**
- [ ] **ステップ 4: 完了報告とクリーンアップ**
