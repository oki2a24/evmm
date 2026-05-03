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

