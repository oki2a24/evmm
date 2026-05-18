import pytest
import os
import openpyxl
from scripts.generate_master_template import TemplateGenerator

def test_generate_from_custom_config(tmp_path):
    """
    JSON 設定に従って、任意のフェーズとカラムを持つエクセルが生成されることを検証。
    """
    output_path = str(tmp_path / "custom_master.xlsx")
    config = {
        "sheet_name": "Custom_WBS",
        "header_row": 2,
        "columns": {
            "common": {
                "id": {"index": 0, "name": "No."},
                "name": {"index": 1, "name": "項目名"}
            },
            "phases": [
                {
                    "name": "独自フェーズA",
                    "mapping": {
                        "plan_start": {"index": 2, "name": "開始"},
                        "plan_effort": {"index": 3, "name": "工数"}
                    }
                }
            ],
            "all": [
                {"name": "No.", "index": 0, "role": "id"},
                {"name": "項目名", "index": 1, "role": "name"},
                {"name": "開始", "index": 2, "role": "plan_start"},
                {"name": "工数", "index": 3, "role": "plan_effort"}
            ]
        }
    }
    
    gen = TemplateGenerator(output_path)
    # 現在の TemplateGenerator は config 引数を持っていないため、ここで失敗するはず
    gen.generate(config=config)
    
    assert os.path.exists(output_path)
    wb = openpyxl.load_workbook(output_path)
    assert "Custom_WBS" in wb.sheetnames
    
    ws = wb["Custom_WBS"]
    assert ws.cell(row=1, column=3).value == "独自フェーズA"
    assert ws.cell(row=2, column=1).value == "No."
    assert ws.cell(row=2, column=2).value == "項目名"
    assert ws.cell(row=2, column=3).value == "開始"
