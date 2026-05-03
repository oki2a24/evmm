import os
import pytest
from openpyxl import load_workbook

from scripts.generate_master_template import TemplateGenerator

# 生成されるテンプレートのパス
TEMPLATE_PATH = "templates/master_template.xlsx"

def test_template_file_exists():
    """
    ステップ 1: テンプレートファイルが物理的に生成されることを検証する。
    """
    if os.path.exists(TEMPLATE_PATH):
        os.remove(TEMPLATE_PATH)

    # 最小限の実装を呼び出す
    generator = TemplateGenerator(TEMPLATE_PATH)
    generator.generate()

    assert os.path.exists(TEMPLATE_PATH), "テンプレートファイルが生成されていません。"

def test_settings_sheet_structure():
    """
    タスク 2: Settings シートの構造を検証する。
    """
    generator = TemplateGenerator(TEMPLATE_PATH)
    generator.generate()
    
    wb = load_workbook(TEMPLATE_PATH)
    assert "Settings" in wb.sheetnames, "Settings シートが存在しません。"
    
    ws = wb["Settings"]
    # 期待される見出しの検証
    assert ws["A1"].value == "プロジェクト基本情報", "A1に見出しがありません。"
    assert ws["A3"].value == "メンバーリスト", "A3に見出しがありません。"

def test_wbs_evm_sheet_structure():
    """
    タスク 3: WBS_EVM シートの 3フェーズ構造と日本語併記を検証する。
    """
    generator = TemplateGenerator(TEMPLATE_PATH)
    generator.generate()
    
    wb = load_workbook(TEMPLATE_PATH)
    assert "WBS_EVM" in wb.sheetnames, "WBS_EVM シートが存在しません。"
    
    ws = wb["WBS_EVM"]
    
    # 3フェーズヘッダーの確認 (結合されていることを想定)
    # F列: 作成, O列: レビュー実施, X列: レビュー後修正 (設計に基づき座標を調整)
    assert ws["F1"].value == "作成", "作成フェーズのヘッダーがありません。"
    assert ws["O1"].value == "レビュー実施", "レビュー実施フェーズのヘッダーがありません。"
    assert ws["X1"].value == "レビュー後修正", "レビュー後修正フェーズのヘッダーがありません。"
    
    # 日本語併記の確認
    # 各フェーズの PV, EV, AC 列を確認
    assert "PV (計画値)" in [ws.cell(row=2, column=c).value for c in range(1, 40)], "PV (計画値) の列が見当たりません。"
    assert "EV (出来高)" in [ws.cell(row=2, column=c).value for c in range(1, 40)], "EV (出来高) の列が見当たりません。"

