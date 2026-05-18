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

def test_dynamic_formula_references(tmp_path):
    """
    動的に配置された列に対して、正しい相対参照の数式が生成されていることを検証。
    """
    output_path = str(tmp_path / "formula_test.xlsx")
    # 工数予定(index:2), 進捗(index:3) の位置に配置
    config = {
        "sheet_name": "WBS_EVM",
        "header_row": 2,
        "columns": {
            "common": {"id": {"index": 0, "name": "ID"}, "name": {"index": 1, "name": "Name"}},
            "phases": [
                {
                    "name": "Phase1",
                    "mapping": {
                        "plan_effort": {"index": 2, "name": "予定工数"},
                        "progress": {"index": 3, "name": "進捗率"},
                        "ev": {"index": 4, "name": "出来高(EV)"}
                    }
                }
            ],
            "all": [
                {"name": "ID", "index": 0, "role": "id"},
                {"name": "Name", "index": 1, "role": "name"},
                {"name": "予定工数", "index": 2, "role": "plan_effort"},
                {"name": "進捗率", "index": 3, "role": "progress"},
                {"name": "出来高(EV)", "index": 4, "role": "ev"}
            ]
        }
    }
    
    gen = TemplateGenerator(output_path)
    gen.generate(config=config)
    
    wb = openpyxl.load_workbook(output_path)
    ws = wb["WBS_EVM"]
    
    # 3行目の EV 数式を検証
    # 予定工数は C列 (3), 進捗率は D列 (4), EVは E列 (5)
    # 期待される数式: =C3 * (D3/100)
    ev_formula = ws.cell(row=3, column=5).value
    assert ev_formula == "=C3*(D3/100)"

def test_cli_integration(tmp_path):
    """
    コマンドライン引数 --config を使用してエクセルが生成されることを検証。
    """
    import json
    import subprocess
    import sys
    
    config = {
        "sheet_name": "CLI_Sheet",
        "header_row": 2,
        "columns": {
            "common": {"id": {"index": 0, "name": "ID"}, "name": {"index": 1, "name": "Name"}},
            "phases": [{"name": "P1", "mapping": {"progress": {"index": 2, "name": "P"}}}],
            "all": [{"name": "ID", "index": 0, "role": "id"}, {"name": "Name", "index": 1, "role": "name"}, {"name": "P", "index": 2, "role": "progress"}]
        }
    }
    config_path = tmp_path / "cli_config.json"
    with open(config_path, "w") as f:
        json.dump(config, f)
        
    output_path = tmp_path / "cli_out.xlsx"
    
    # スクリプトの実行 (python scripts/generate_master_template.py --config ... --output ...)
    cmd = [
        sys.executable, 
        "scripts/generate_master_template.py", 
        "--config", str(config_path), 
        "--output", str(output_path)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert os.path.exists(output_path)
    
    wb = openpyxl.load_workbook(output_path)
    assert "CLI_Sheet" in wb.sheetnames
