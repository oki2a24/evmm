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

def test_check_workload_overload():
    """
    稼働日数に対して予定工数が超過している（過負荷）を検知できるか検証する。
    """
    data = {
        "機能ID": ["F4"],
        "機能名称": ["Overload Test"],
        "開始日予定": [datetime.date(2026, 4, 1)], # 水
        "終了日予定": [datetime.date(2026, 4, 2)], # 木 (稼働日2日)
        "工数予定": [3.0], # 2日に対して3人日はオーバー
        "担当メンバー": ["A"],
        "チームリーダー": ["L"]
    }
    df = pd.DataFrame(data)
    
    checker = WBSIntegrityChecker("dummy.xlsx")
    errors = checker.check_dataframe(df)
    
    assert any("工数予定(3.0)が稼働日数(2)を超過" in e["message"] for e in errors)

def test_check_phase_sequence():
    """
    作成・レビュー・修正のフェーズ間順序の矛盾を検知できるか検証する。
    """
    data = {
        "機能ID": ["F5"],
        "機能名称": ["Sequence Test"],
        "開始日予定": [datetime.date(2026, 4, 1)],
        "終了日予定": [datetime.date(2026, 4, 2)],
        "工数予定": [1],
        "担当メンバー": ["A"],
        "チームリーダー": ["L"],
        "開始日予定.1": [datetime.date(2026, 4, 1)], # レビュー開始が作成終了より前
        "終了日予定.1": [datetime.date(2026, 4, 1)],
        "工数予定.1": [1]
    }
    df = pd.DataFrame(data)
    
    checker = WBSIntegrityChecker("dummy.xlsx")
    errors = checker.check_dataframe(df)
    
    # メッセージに期待するキーワードが含まれているか確認
    assert any("作成フェーズの終了日" in e["message"] and "レビューフェーズが開始" in e["message"] for e in errors)

def test_check_actual_integrity():
    """
    実績入力（進捗率100%時の終了日実績、工数と日付の矛盾）を検知できるか検証する。
    """
    data = {
        "機能ID": ["F6"],
        "機能名称": ["Actual Test"],
        "開始日予定": [datetime.date(2026, 4, 1)],
        "終了日予定": [datetime.date(2026, 4, 2)],
        "工数予定": [1],
        "担当メンバー": ["A"],
        "チームリーダー": ["L"],
        "進捗率(%)": [100.0],
        "終了日実績": [pd.NaT], # 100%なのに終了日がない
        "工数実績": [1.0]
    }
    df = pd.DataFrame(data)
    
    checker = WBSIntegrityChecker("dummy.xlsx")
    errors = checker.check_dataframe(df)
    
    assert any("進捗率100%ですが、終了日実績が未入力" in e["message"] for e in errors)
