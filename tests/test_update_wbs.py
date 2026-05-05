import pytest
import openpyxl
import pandas as pd
import shutil
import os
from datetime import datetime
from scripts.update_wbs import update_wbs_logic

# テスト用のダミーWBSパス
TEST_EXCEL = "tests/test_data_update.xlsx"
TEMPLATE_EXCEL = "projects/test_project/hoge_wbs_evm.xlsx"

@pytest.fixture
def setup_excel():
    """テスト実行前にテンプレートをコピーして環境を整える"""
    if not os.path.exists("tests"):
        os.makedirs("tests")
    shutil.copy(TEMPLATE_EXCEL, TEST_EXCEL)
    yield TEST_EXCEL
    # テスト後に削除
    if os.path.exists(TEST_EXCEL):
        os.remove(TEST_EXCEL)

def test_update_wbs_basic(setup_excel):
    """
    基本機能のテスト: F1の『作成』フェーズを工数3hで更新する。
    """
    # 実行前の値を確認（F1, 作成, 工数実績は通常最初は空か別の値）
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb['WBS_EVM']
    # F1は3行目(skiprows=1を考慮するとdfの0行目、Excelの3行目)
    # 作成フェーズの工数実績はK列(11列目)
    original_effort = ws.cell(row=3, column=11).value
    wb.close()

    # 更新実行
    new_effort = 3.0
    update_wbs_logic(setup_excel, func_id="F1", phase="作成", effort=new_effort)

    # 実行後の値を確認
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb['WBS_EVM']
    updated_effort = ws.cell(row=3, column=11).value
    updated_progress = ws.cell(row=3, column=12).value # 進捗率 L列
    wb.close()

    assert updated_effort == new_effort
    assert updated_progress == 100 # 完了報告なので100%を期待

def test_update_wbs_stateful(setup_excel):
    """
    ステートフル更新のテスト: 
    1. F1の『作成』を100%にする。
    2. phaseを指定せずに呼び出した場合、自動的に次の『レビュー実施』が更新されることを確認する。
    """
    # 1. まず『作成』を完了させる
    update_wbs_logic(setup_excel, func_id="F1", phase="作成", effort=2.0)
    
    # 2. 次に phase を None (または自動判別を期待する値) で呼び出す
    # ※ 現在のロジックは phase 指定が必須なので、ここで失敗するはず
    new_effort_review = 1.5
    update_wbs_logic(setup_excel, func_id="F1", phase=None, effort=new_effort_review)

    # 3. 『レビュー実施』の列(occurrence=1)が更新されているか確認
    # レビュー実施フェーズの工数実績は W列(23列目)
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb['WBS_EVM']
    updated_effort_review = ws.cell(row=3, column=23).value
    wb.close()

    assert updated_effort_review == new_effort_review

def test_update_wbs_atomicity(setup_excel):
    """
    原子性のテスト: 
    整合性エラーが発生するような状態で更新を試みた際、
    例外が投げられ、かつ元のファイルが書き換わっていないことを確認する。
    """
    # 1. 事前にファイルを『壊れた』状態（開始日予定 > 終了日予定）にする
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb['WBS_EVM']
    # F1の開始日予定(F列=6)を終了日予定(G列=7)より後にする
    ws.cell(row=3, column=6, value=datetime(2026, 12, 31))
    ws.cell(row=3, column=7, value=datetime(2026, 1, 1))
    
    # 元の工数実績(K列=11)を記録
    original_effort = ws.cell(row=3, column=11).value
    wb.save(setup_excel)
    wb.close()

    # 2. 更新を試みる。内部で整合性チェックが走り、ValueError が出るはず。
    with pytest.raises(ValueError) as excinfo:
        update_wbs_logic(setup_excel, func_id="F1", phase="作成", effort=99.0)
    
    assert "整合性チェックエラー" in str(excinfo.value)

    # 3. 重要: ファイルの工数実績が書き換わっていないことを確認
    wb = openpyxl.load_workbook(setup_excel)
    ws = wb['WBS_EVM']
    final_effort = ws.cell(row=3, column=11).value
    wb.close()

    # 現状のバグ（原子性なし）では、保存後にチェックするため、99.0 に書き換わってしまっている
    assert final_effort == original_effort, f"エラーが発生したのにファイルが更新されてしまいました（{final_effort} != {original_effort}）"
