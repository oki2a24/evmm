# 将来予測 (EAC/ETC) 実装計画

> **AIエージェントへの指示:** REQUIRED SUB-SKILL: 本計画をタスクごとに実装するには、`subagent-driven-development` スキル（推奨）を `activate_skill` で起動して使用してください。
> 本計画は **TDD (Red-Green-Refactor)** に基づき、PMBOK 理論に基づいた正確な将来予測ロジックを段階的に構築する。

**目標:** 
1. `analyze_evm.py` への BAC, ETC, EAC, VAC 計算ロジックの実装。
2. 楽観・現実・慎重の 3 シナリオによる将来予測機能の提供。
3. AI アドバイザー（evm_analyst）によるインテリジェントな予測解釈。

**アーキテクチャ:** 
- **Calculator (Python)**: `scripts/analyze_evm.py`。PMBOK 公式に基づく計算エンジン。
- **Persona (Sub-Agent)**: `evm_analyst`。予測値を基にリスクと対策を語る専門家。

---

### タスク 1: BAC (完成時総予算) 算出の実装 (TDD)

**ファイル:**
- 変更: `scripts/analyze_evm.py`
- 変更: `tests/test_analyze_evm.py`

- [ ] **ステップ 1: BAC 算出の RED テスト作成**
  - WBS全体の「工数予定」が、基準日に左右されず正しく合計されることを検証。
- [ ] **ステップ 2: `EVMAnalyst` への BAC 計算ロジック追加 (GREEN)**
- [ ] **ステップ 3: テスト実行と GREEN 確認**
  - `pytest tests/test_analyze_evm.py`
- [ ] **ステップ 4: コミット**

### タスク 2: 3 シナリオ予測ロジックの実装 (TDD)

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: 予測計算の RED テスト作成**
  - 既知の CPI/SPI 値を用い、楽観・現実・慎重の 3 パターンの EAC/ETC が公式通り算出されるか検証。
- [ ] **ステップ 2: `EVMAnalyst` への予測計算メソッド実装 (GREEN)**
  - `calculate_forecasts()` 等のメソッドを新設。
- [ ] **ステップ 3: テスト実行と GREEN 確認**
- [ ] **ステップ 4: コミット**

### タスク 3: JSON サマリーの拡張と統合

**ファイル:**
- 変更: `scripts/analyze_evm.py`

- [ ] **ステップ 1: `get_summary_json` 拡張の RED テスト作成**
  - `forecasting` セクションが含まれ、各シナリオの数値が正しい構造で出力されるか検証。
- [ ] **ステップ 2: `get_summary_json` への統合実装 (GREEN)**
- [ ] **ステップ 3: 実ファイルによる結合テスト**
  - `projects/test_project/hoge_wbs_evm.xlsx` 等を用いて、一連の分析フローを検証。
- [ ] **ステップ 4: コミット**

### タスク 4: AI アドバイザー (evm_analyst) の強化

**ファイル:**
- 変更: `.gemini/agents/evm_analyst.md`

- [ ] **ステップ 1: `writing-skills` スキルを用いた指示の洗練**
  - `writing-skills` スキルを起動し、エージェント定義を最新の予測ロジックに合わせて体系的に更新する。
  - EAC/VAC の数値を基に、「予算内に収めるために必要な目標 CPI (TCPI)」を逆算して提示する等の思考ロジックを追加。
- [ ] **ステップ 2: 教育的フィードバックの強化**
  - 予測シナリオ（楽観 vs 慎重）の差が何を意味するかをユーザーに解説する指示を強化。
- [ ] **ステップ 3: 動作検証**
  - 予測データを含む JSON を与え、期待通りのインテリジェントなアドバイスが生成されるか確認。
- [ ] **ステップ 4: コミット**

---

## 完了の定義 (DoD)
- [ ] `analyze_evm.py` が 3 つのシナリオで正確な EAC/ETC/VAC を算出できる。
- [ ] 全ての新規ユニットテストおよび既存テストがパスする。
- [ ] AI アナリスト（evm_analyst）が予測値を基に具体的なリスク対策を提案できる。
- [ ] `code-reviewer` による最終レビューを通過する。
