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

def test_save_and_load_config_integrity(tmp_path):
    """
    infer_structure で生成した設定を保存し、再読み込みした際に
    get_column_index などの挙動が完全に一致することを検証する。
    """
    df = pd.DataFrame([
        ["作成", None, None, None, None],
        ["機能ID", "機能名称", "開始日予定", "進捗率", "備考"],
        ["F1", "Test", "2026-05-01", 0, "テスト備考"]
    ])
    file_path = tmp_path / "integrity_test.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    config = manager.infer_structure()
    manager.config = config
    manager.save_config()
    
    # 再読み込み
    new_manager = WBSConfigManager(str(file_path))
    loaded_config = new_manager.load_or_infer(interactive=False)
    
    # 役割ベースでのインデックス取得が一致することを確認
    assert new_manager.get_column_index("id") == 0
    assert new_manager.get_column_index("name") == 1
    assert new_manager.get_column_index("plan_start", phase_idx=0) == 2
    assert new_manager.get_column_index("progress", phase_idx=0) == 3

    # 未知の列（備考など）も保持されていることを期待（ここが現在のRED要件）
    # 現在の infer_structure は役割にない列を無視しているため
    assert "備考" in [c["name"] for c in loaded_config["columns"]["all"]]

def test_infer_with_aliases_and_insertion(tmp_path):
    """
    実証テストで遭遇した『意地悪な構造』に対する推論テスト。
    1. A列の前に空列を挿入（全体がずれる）
    2. カラム名に別名（作業開始日、期限）を使用
    3. フェーズを2つに減らす
    """
    # 1列目に空列（None）を入れたデータ
    # ID, 名称は共通。
    # 「作成」フェーズ（作業開始日、期限）は 3, 4列目。
    # 「レビュー」フェーズ（進捗、実績）は 5, 6列目。
    df = pd.DataFrame([
        [None, None, None, "作成", None, "レビュー", None],
        [None, "機能ID", "機能名称", "作業開始日", "期限", "進捗率", "実績工数"],
        [None, "F1", "Test", "2026-05-01", "2026-05-05", 0, 0]
    ])

    file_path = tmp_path / "messy_test.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    config = manager.infer_structure()

    # ID列のインデックスが 1 (B列) になっているはず
    assert config["columns"]["common"]["id"]["index"] == 1
    # 「作業開始日」が plan_start として認識されているはず
    assert config["columns"]["phases"][0]["mapping"]["plan_start"]["name"] == "作業開始日"
    # 「期限」が plan_end として認識されているはず
    assert config["columns"]["phases"][0]["mapping"]["plan_end"]["name"] == "期限"
    # フェーズ数が2つであることを認識
    assert len(config["columns"]["phases"]) == 2

def test_missing_config_fails_gracefully(tmp_path):
    """
    必須の ID 列が存在しない場合に、分かりやすいエラーを投げるか検証。
    """
    df = pd.DataFrame([
        ["名前しかないシート"],
        ["太郎"],
        ["次郎"]
    ])
    file_path = tmp_path / "no_id.xlsx"
    df.to_excel(file_path, index=False, header=False, sheet_name="WBS_EVM")

    manager = WBSConfigManager(str(file_path))
    with pytest.raises(ValueError, match="Could not find header row"):
        manager.infer_structure()


