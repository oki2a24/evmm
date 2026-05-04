import pytest
import pandas as pd
import datetime
from scripts.check_wbs_integrity import WBSIntegrityChecker

def test_check_date_reversal():
    """
    開始日が終了日より後になっている不整合を検知できるか検証する。
    """
    data = {
        "機能ID": ["F1", "F2"],
        "機能名称": ["Test1", "Test2"],
        "開始日予定": [datetime.date(2026, 4, 1), datetime.date(2026, 4, 5)],
        "終了日予定": [datetime.date(2026, 4, 3), datetime.date(2026, 4, 1)], # F2が逆転
        "工数予定": [2, 1],
        "担当メンバー": ["A", "B"],
        "チームリーダー": ["L", "L"]
    }
    df = pd.DataFrame(data)
    
    checker = WBSIntegrityChecker("dummy.xlsx")
    errors = checker.check_dataframe(df)
    
    # F2 (index 1) にエラーがあるはず
    f2_errors = [e for e in errors if e["index"] == 1]
    assert any("開始日予定が終了日予定より後" in e["message"] for e in f2_errors)
    
    # F1 (index 0) にはエラーがないはず
    f1_errors = [e for e in errors if e["index"] == 0]
    assert not any("開始日予定が終了日予定より後" in e["message"] for e in f1_errors)

def test_missing_required_dates():
    """
    予定日が欠落している不整合を検知できるか検証する。
    """
    data = {
        "機能ID": ["F3"],
        "機能名称": ["Test3"],
        "開始日予定": [pd.NaT],
        "終了日予定": [datetime.date(2026, 4, 1)],
        "工数予定": [1],
        "担当メンバー": ["A"],
        "チームリーダー": ["L"]
    }
    df = pd.DataFrame(data)
    
    checker = WBSIntegrityChecker("dummy.xlsx")
    errors = checker.check_dataframe(df)
    
    assert any("開始日予定が未入力" in e["message"] for e in errors)
