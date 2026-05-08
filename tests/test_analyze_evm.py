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

def test_calculate_bac(tmp_path):
    """
    BAC (完成時総予算) 算出のテスト。
    基準日に左右されず、WBS全体の「工数予定」が正しく合計されることを検証。
    """
    import shutil
    import pandas as pd
    from scripts.analyze_evm import EVMAnalyst

    # テスト用テンプレートを一時ディレクトリにコピー
    template_path = "tests/data/wbs_template_for_testing.xlsx"
    test_excel = tmp_path / "test_bac.xlsx"
    shutil.copy(template_path, test_excel)

    # テストデータの準備: 工数予定(4列目)に値をセット
    # テンプレート構造に基づき、工数予定は 8列目(H), 20列目(T), 32列目(AF) 等
    import openpyxl
    wb = openpyxl.load_workbook(str(test_excel))
    ws = wb['WBS_EVM']
    # 3行目: 作成(8列目)=10, レビュー(20列目)=2, 修正(32列目)=1 -> 計13
    ws.cell(row=3, column=8, value=10.0)
    ws.cell(row=3, column=20, value=2.0)
    ws.cell(row=3, column=32, value=1.0)
    # 4行目: 作成(8列目)=5 -> 計18
    ws.cell(row=4, column=8, value=5.0)
    wb.save(str(test_excel))

    analyst = EVMAnalyst(file_path=str(test_excel))
    
    # まだメソッドが実装されていないため、ここで AttributeError またはエラーを期待 (RED)
    # 内部的に DataFrame をロードして集計するロジックが必要
    df = pd.read_excel(test_excel, sheet_name='WBS_EVM', header=1)
    bac = analyst.calculate_bac(df)
    
    assert bac == 18.0

def test_calculate_forecasts():
    """
    将来予測（3シナリオ）の計算テスト。
    PMBOK公式に基づき、EAC, ETC, VACが正しく算出されるか検証。
    """
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    
    # テストデータ
    bac = 100.0
    ev = 40.0
    ac = 50.0 # CPI = 40/50 = 0.8 (コスト超過)
    pv = 50.0 # SPI = 40/50 = 0.8 (スケジュール遅延)
    
    forecasts = analyst.calculate_forecasts(bac, ev, ac, pv)
    
    # 1. 現実的 (Realistic): EAC = BAC / CPI = 100 / 0.8 = 125.0
    # ETC = EAC - AC = 125.0 - 50.0 = 75.0
    # VAC = BAC - EAC = 100.0 - 125.0 = -25.0
    res = forecasts["realistic"]
    assert res["eac"] == 125.0
    assert res["etc"] == 75.0
    assert res["vac"] == -25.0
    
    # 2. 楽観的 (Optimistic): EAC = AC + (BAC - EV) = 50.0 + (100.0 - 40.0) = 110.0
    # ETC = 60.0, VAC = -10.0
    opt = forecasts["optimistic"]
    assert opt["eac"] == 110.0
    assert opt["vac"] == -10.0
    
    # 3. 慎重 (Pessimistic): EAC = AC + [(BAC - EV) / (CPI * SPI)]
    # EAC = 50.0 + [(100.0 - 40.0) / (0.8 * 0.8)] = 50.0 + [60.0 / 0.64] = 50.0 + 93.75 = 143.75
    pes = forecasts["pessimistic"]
    assert pes["eac"] == 143.75
    assert pes["vac"] == -43.75

def test_calculate_forecasts_zero_efficiency():
    """効率が0（未着手または極端な効率低下）の場合のガードレール検証"""
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    
    # EV=0 (CPI=0) の場合、現実的/慎重モデルは無限大になるが、
    # 実装上は極めて大きな値または特定の警告値を期待。
    # ここではゼロ除算でクラッシュしないことを検証。
    forecasts = analyst.calculate_forecasts(bac=100.0, ev=0.0, ac=10.0, pv=10.0)
    assert forecasts["realistic"]["eac"] >= 1000.0 # ある程度の安全値または上限値
    assert forecasts["optimistic"]["eac"] == 110.0 # 楽観的は影響を受けない
