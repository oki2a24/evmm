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

