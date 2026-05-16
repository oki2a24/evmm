import pytest
import pandas as pd
import os
from scripts.wbs_config_manager import WBSConfigManager

def test_infer_and_save_load(tmp_path):
    # テスト用のエクセル作成
    df = pd.DataFrame([
        ["作成", None, None, "レビュー", None, None],
        ["機能ID", "機能名称", "開始日予定", "終了日予定", "進捗率(%)", "AC"],
        ["F1", "Test", "2026-05-01", "2026-05-05", 0, 0]
    ])
    file_path = tmp_path / "test_wbs.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    # インタラクティブなしで実行
    config = manager.load_or_infer(interactive=False)
    manager.save_config()

    # インデックスの確認 (0-based)
    assert config["columns"]["common"]["id"]["index"] == 0
    assert config["columns"]["phases"][0]["mapping"]["plan_start"]["index"] == 2
    
    # JSONが作成されているか
    assert os.path.exists(manager.config_path)

    # 再読み込み
    new_manager = WBSConfigManager(str(file_path))
    loaded_config = new_manager.load_or_infer(interactive=False)
    assert loaded_config["header_row"] == 2
    assert loaded_config["columns"]["common"]["id"]["name"] == "機能ID"

def test_get_column_index(tmp_path):
    df = pd.DataFrame([
        ["作成", None],
        ["機能ID", "開始日予定"],
        ["F1", "2026-05-01"]
    ])
    file_path = tmp_path / "test_wbs_idx.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    assert manager.get_column_index("id") == 0
    assert manager.get_column_index("plan_start", phase_idx=0) == 1

