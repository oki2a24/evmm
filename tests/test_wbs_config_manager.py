import pytest
import pandas as pd
import os
from scripts.wbs_config_manager import WBSConfigManager

def test_infer_header_and_columns(tmp_path):
    # テスト用のエクセル作成
    df = pd.DataFrame([
        ["作成", None, None, "レビュー", None, None],
        ["機能ID", "機能名称", "開始日予定", "開始日予定", "進捗率(%)", "進捗率(%)"],
        ["F1", "Test", "2026-05-01", "2026-05-05", 0, 0]
    ])
    file_path = tmp_path / "test_wbs.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    config = manager.infer_structure()

    assert config["header_row"] == 2
    assert config["columns"]["common"]["id"] == "機能ID"
    assert len(config["columns"]["phases"]) >= 2
    assert config["columns"]["phases"][0]["name"] == "作成"
    assert config["columns"]["phases"][0]["mapping"]["plan_start"] == "開始日予定"
    assert config["columns"]["phases"][1]["name"] == "レビュー"
