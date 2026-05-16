import pytest
import pandas as pd
import datetime
import os
from scripts.check_wbs_integrity import WBSIntegrityChecker

@pytest.fixture
def temp_checker(tmp_path):
    """
    テストデータに合わせたマッピングを持つチェッカーを作成する。

    【テスト設計の重要事項】
    WBSConfigManager はフェーズ行（1行目）とヘッダー行（2行目）の相対関係で
    マッピングを推論します。テストで期待通りのインデックスを取得するためには、
    この2行の構造を正確に模倣したエクセルファイルを生成する必要があります。
    """
    def _create(df_headers):

        file_path = tmp_path / "test_check.xlsx"
        df = pd.DataFrame([
            df_headers[0],
            df_headers[1],
            ["F1"] + [None] * (len(df_headers[1]) - 1)
        ])
        df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")
        return WBSIntegrityChecker(str(file_path), interactive=False)
    return _create

def test_check_date_reversal(temp_checker):
    """
    開始日が終了日より後になっている不整合を検知できるか検証する。
    """
    headers = [
        [None, None, "作成", None, None, None, None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "担当メンバー", "チームリーダー"]
    ]
    checker = temp_checker(headers)
    
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
    errors = checker.check_dataframe(df)
    
    f2_errors = [e for e in errors if e["index"] == 1]
    assert any("開始日予定が終了日予定より後" in e["message"] for e in f2_errors)

def test_missing_required_dates(temp_checker):
    """
    予定日が欠落している不整合を検知できるか検証する。
    """
    headers = [
        [None, None, "作成", None, None, None, None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "担当メンバー", "チームリーダー"]
    ]
    checker = temp_checker(headers)
    
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
    errors = checker.check_dataframe(df)
    
    assert any("開始日予定が未入力" in e["message"] for e in errors)

def test_check_workload_overload(temp_checker):
    """
    稼働日数に対して予定工数が超過している（過負荷）を検知できるか検証する。
    """
    headers = [
        [None, None, "作成", None, None, None, None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "担当メンバー", "チームリーダー"]
    ]
    checker = temp_checker(headers)
    
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
    errors = checker.check_dataframe(df)
    
    assert any("工数予定(3.0)が稼働日数(2)を超過" in e["message"] for e in errors)

def test_check_phase_sequence(temp_checker):
    """
    作成・レビュー・修正のフェーズ間順序の矛盾を検知できるか検証する。
    """
    headers = [
        [None, None, "作成", None, None, "レビュー実施", None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "開始日予定", "終了日予定"]
    ]
    checker = temp_checker(headers)
    
    data = {
        "機能ID": ["F5"],
        "機能名称": ["Sequence Test"],
        "開始日予定": [datetime.date(2026, 4, 1)],
        "終了日予定": [datetime.date(2026, 4, 2)],
        "工数予定": [1],
        "開始日予定.1": [datetime.date(2026, 4, 1)], # レビュー開始が作成終了より前
        "終了日予定.1": [datetime.date(2026, 4, 1)]
    }
    df = pd.DataFrame(data)
    df.columns = ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "開始日予定.1", "終了日予定.1"]
    
    errors = checker.check_dataframe(df)
    assert any("前フェーズの終了日" in e["message"] and "レビュー実施フェーズが開始" in e["message"] for e in errors)


def test_check_actual_integrity(temp_checker):
    """
    実績入力（進捗率100%時の終了日実績）を検知できるか検証する。
    """
    headers = [
        [None, None, "作成", None, None, None, None, None, None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "工数予定", "進捗率(%)", "開始日実績", "終了日実績", "工数実績"]
    ]
    checker = temp_checker(headers)
    
    data = {
        "機能ID": ["F6"],
        "機能名称": ["Actual Test"],
        "開始日予定": [datetime.date(2026, 4, 1)],
        "終了日予定": [datetime.date(2026, 4, 2)],
        "工数予定": [1],
        "進捗率(%)": [100.0],
        "開始日実績": [pd.NaT],
        "終了日実績": [pd.NaT], # 100%なのに終了日がない
        "工数実績": [1.0]
    }
    df = pd.DataFrame(data)
    errors = checker.check_dataframe(df)
    
    assert any("進捗率100%ですが、終了日実績が未入力" in e["message"] for e in errors)
