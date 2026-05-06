# EVM 分析 & AI アナリスト (EVMAnalyst) 実装計画 V2

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: `subagent-driven-development` スキル（推奨）を `activate_skill` で起動して使用してください。
> 本計画は **TDD (Red-Green-Refactor)** を核とし、計算の正確性（Python）と解釈の高度化（カスタムサブエージェント）を両立させる「三位一体」の実装プランである。

**目標:** 
1. Python スクリプトによる正確な EVM 計算と Excel 反映。
2. EVM 分析に特化したカスタムサブエージェント `evm_analyst` の構築。
3. `/analyze-evm` スキルによる、分析フローの規律化（整合性チェックの義務化）。

**アーキテクチャ:** 
- **Calculator (Python)**: `scripts/analyze_evm.py`。稼働日ベースの正確な計算機。
- **Persona (Sub-Agent)**: `evm_analyst`。PMBOK 理論に基づき診断する専門家。
- **Controller (Skill)**: `.gemini/skills/analyze-evm/SKILL.md`。一連のフローを制御。

---

### タスク 1: 計算エンジン `analyze_evm.py` の基本実装 (TDD)

**ファイル:**
- 作成: `scripts/analyze_evm.py`, `tests/test_analyze_evm.py`

- [ ] **ステップ 1: PV 算出の RED テスト作成**
  - 基準日に基づき、`(基準日 - 開始予定日) / (終了予定日 - 開始予定日)` が稼働日ベースで正しく線形按分されることを検証。
- [ ] **ステップ 2: `EVMAnalyst.calculate_pv` の GREEN 実装**
  - `scripts/utils.py` の `get_working_days` を使用。
- [ ] **ステップ 3: EV/AC 算出の RED テスト作成**
  - `EV = 工数予定 * 進捗率`, `AC = 工数実績の合算` を検証。
- [ ] **ステップ 4: `calculate_ev`, `calculate_ac` の GREEN 実装**
- [ ] **ステップ 5: code-reviewer 召喚 & コミット**

### タスク 2: 整合性チェック統合と Excel 外科的更新

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: 整合性チェック連携の RED テスト作成**
  - エラーがある場合に分析を中断し、警告を返す挙動を検証。
- [ ] **ステップ 2: `WBSIntegrityChecker` 連携の実装 (GREEN)**
- [ ] **ステップ 3: Excel 書き込みの RED テスト作成**
  - `openpyxl` を使い、既存の書式や計算式を壊さず PV/EV/AC 列のみを上書きできるか検証。
- [ ] **ステップ 4: `write_results_to_excel` の実装 (GREEN)**
- [ ] **ステップ 5: code-reviewer 召喚 & コミット**

### タスク 3: カスタムサブエージェント `evm_analyst` の構築

**ファイル:**
- 作成: `.gemini/agents/evm_analyst.md`

- [ ] **ステップ 1: ペルソナ定義の作成**
  - PMBOK 理論を背景に持ち、CPI/SPI の数値から「プロジェクトの真の危機」を読み解く指示セットを定義。
- [ ] **ステップ 2: 入出力インターフェースの設計**
  - Calculator の計算結果（JSON形式）を解釈し、日本語でアドバイスを生成するプロンプトの調整。
- [ ] **ステップ 3: テスト用診断テスト**
  - 複数の遅延パターン（効率低下、リソース過負荷等）を与え、適切なアドバイスが出るか検証。
- [ ] **ステップ 4: コミット**

### タスク 4: `/analyze-evm` スキルの構築

**ファイル:**
- 作成: `.gemini/skills/analyze-evm/SKILL.md`

- [ ] **ステップ 1: スキル定義の作成**
  - 分析実行時の「思考プロセス」を定義（整合性チェック義務化、基準日の確認など）。
- [ ] **ステップ 2: フロー制御の実装**
  - `analyze_evm.py` の実行結果を `evm_analyst` サブエージェントに渡す連携パスを記述。
- [ ] **ステップ 3: コミット**

### タスク 5: 統合検証と最終調整 (TDD継続)

- [ ] **ステップ 1: 実データ（4/3時点）への適用テスト**
  - `projects/test_project/hoge_wbs_evm.xlsx` で一括実行。
- [ ] **ステップ 2: 不備の修正**
  - 期待と異なる結果が出た場合、まずテストコードを修正してからロジックを直す。
- [ ] **ステップ 3: code-reviewer による最終承認**
- [ ] **ステップ 4: 完了報告とクリーンアップ**
