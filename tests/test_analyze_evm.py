import pytest
from datetime import date
from scripts.analyze_evm import EVMAnalyst

"""
EVM分析エンジンのテストコード

【テストの目的】
本テストは、EVMにおけるPV（Planned Value: 計画値）の計算が、
PMBOK理論に基づき、かつ「稼働日（土日祝日を除く）」を考慮して
正しく線形按分されることを物理的に証明します。

【設計上の注意】
- 稼働日計算には既存の scripts/utils.py のロジックを利用することを前提とします。
- 基準日（Status Date）がタスク期間外（開始前、終了後）の場合の境界値も検証します。
"""

def test_calculate_pv_linear():
    """
    PVの線形按分計算のテスト（正常系：期間中）
    
    【シナリオ】
    - 開始予定: 2026-04-01 (水)
    - 終了予定: 2026-04-14 (火) 
      -> 土日を除くと 10稼働日 (04/01-03, 04/06-10, 04/13-14)
    - 工数予定: 10.0 人日
    - 基準日: 2026-04-07 (火) 
      -> 04/01から数えて 5稼働日経過
    【期待値】
    - PV = (経過稼働日 / 総稼働日) * 工数予定 = (5 / 10) * 10.0 = 5.0
    """
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 4, 7))
    
    start_date = date(2026, 4, 1)
    end_date = date(2026, 4, 14)
    planned_effort = 10.0
    
    pv = analyst.calculate_pv(start_date, end_date, planned_effort)
    assert pv == 5.0

def test_calculate_pv_before_start():
    """基準日が開始日より前の場合、PVは 0.0 であるべき"""
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 3, 31))
    pv = analyst.calculate_pv(date(2026, 4, 1), date(2026, 4, 10), 10.0)
    assert pv == 0.0

def test_calculate_pv_after_end():
    """基準日が終了日より後の場合、PVは全額（工数予定）であるべき"""
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 4, 15))
    pv = analyst.calculate_pv(date(2026, 4, 1), date(2026, 4, 10), 10.0)
    assert pv == 10.0

def test_calculate_ev():
    """EV = 工数予定 * 進捗率"""
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    ev = analyst.calculate_ev(10.0, 50.0)
    assert ev == 5.0
    
    ev_zero = analyst.calculate_ev(10.0, 0.0)
    assert ev_zero == 0.0

def test_calculate_ac():
    """AC = 工数実績の値そのもの"""
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    ac = analyst.calculate_ac(8.5)
    assert ac == 8.5
    
    ac_none = analyst.calculate_ac(None)
    assert ac_none == 0.0

def test_run_with_integrity_error(mocker):
    """
    整合性エラーがある場合に分析を中断することを検証。
    【目的】壊れたデータに基づいた誤った分析結果を出力することを防ぐ。
    """
    from scripts.check_wbs_integrity import WBSIntegrityChecker
    
    # WBSIntegrityChecker.check_dataframe がエラーを返すようにモック
    mocker.patch.object(WBSIntegrityChecker, 'load_wbs', return_value=None)
    mocker.patch.object(WBSIntegrityChecker, 'check_dataframe', return_value=[{"index": 0, "message": "Error"}])
    mocker.patch.object(WBSIntegrityChecker, 'write_results_to_excel', return_value=None)
    
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    
    # 整合性エラーがある場合、ValueError を送出することを期待
    with pytest.raises(ValueError, match="整合性チェックでエラーが検出されました"):
        analyst.run()

def test_write_results_to_excel(tmp_path):
    """
    Excelへの書き出しが正しく行われるかを検証。
    【重要】既存の書式を壊さず、PV/EV/AC 列のみが更新されること。
    """
    import shutil
    import openpyxl
    from scripts.analyze_evm import EVMAnalyst
    
    # テスト用テンプレートを一時ディレクトリにコピー
    template_path = "tests/data/wbs_template_for_testing.xlsx"
    test_excel = tmp_path / "test_analyze.xlsx"
    shutil.copy(template_path, test_excel)
    
    analyst = EVMAnalyst(file_path=str(test_excel), status_date=date(2026, 4, 7))
    
    # ダミーの計算結果
    # 行3 (index 0) に PV=5.0, EV=2.0, AC=1.5 を書き込む想定
    # カラム位置は通常 WBS の構造に依存
    results = [
        {"row": 3, "pv": 5.0, "ev": 2.0, "ac": 1.5}
    ]
    
    analyst.write_results_to_excel(results)
    
    # 書き込み後のファイルを再読込して検証
    wb = openpyxl.load_workbook(str(test_excel))
    ws = wb['WBS_EVM']
    
    # カラム名を特定 (PV (計画値), EV (出来高), AC (実績コスト))
    headers = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    idx_pv = headers.index("PV (計画値)") + 1
    idx_ev = headers.index("EV (出来高)") + 1
    idx_ac = headers.index("AC (実績コスト)") + 1
    
    assert ws.cell(row=3, column=idx_pv).value == 5.0
    assert ws.cell(row=3, column=idx_ev).value == 2.0
    assert ws.cell(row=3, column=idx_ac).value == 1.5

def test_run_does_not_overwrite_formulas(tmp_path):
    """
    run() を実行しても WBS シートの数式が数値で上書きされないことを検証。
    【重要】エクセルの柔軟性（数式による自動計算）を維持するため。
    """
    import shutil
    import openpyxl
    from scripts.analyze_evm import EVMAnalyst
    
    # テスト用テンプレートを一時ディレクトリにコピー
    template_path = "tests/data/wbs_template_for_testing.xlsx"
    test_excel = tmp_path / "test_formula_protection.xlsx"
    shutil.copy(template_path, test_excel)
    
    # 実行前の数式を確認
    wb_pre = openpyxl.load_workbook(str(test_excel))
    ws_pre = wb_pre['WBS_EVM']
    # PV列（例：13列目）の3行目に数式が入っていることを確認
    formula_pre = ws_pre.cell(row=3, column=13).value
    assert isinstance(formula_pre, str) and formula_pre.startswith("=")
    
    # 分析を実行 (デフォルト設定)
    analyst = EVMAnalyst(file_path=str(test_excel), status_date=date(2026, 4, 7))
    analyst.run()
    
    # 実行後の数式を確認
    wb_post = openpyxl.load_workbook(str(test_excel))
    ws_post = wb_post['WBS_EVM']
    formula_post = ws_post.cell(row=3, column=13).value
    
    # 数式が維持されている（文字列かつ '=' で始まる）ことを期待
    # 現在の実装では数値に書き換わるため、ここで失敗するはず
    assert isinstance(formula_post, str), f"数式が数値に上書きされました: {formula_post}"
    assert formula_post.startswith("="), f"数式が失われました: {formula_post}"
