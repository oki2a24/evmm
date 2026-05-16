import pandas as pd
import openpyxl
import re
import os
import json

class WBSConfigManager:
    """
    WBSエクセルの構造を解析し、役割ベースのカラムマッピングを管理するクラス。
    """
    
    DEFAULT_ROLES = {
        "id": [r"機能ID", r"ID"],
        "name": [r"機能名称", r"タスク名", r"名称"],
        "plan_start": [r"開始日予定", r"予定開始日", r"開始予定"],
        "plan_end": [r"終了日予定", r"予定終了日", r"終了予定"],
        "plan_effort": [r"工数予定", r"予定工数", r"見積工数"],
        "actual_start": [r"開始日実績", r"実績開始日", r"開始実績"],
        "actual_end": [r"終了日実績", r"実績終了日", r"終了実績"],
        "actual_effort": [r"工数実績", r"実績工数"],
        "progress": [r"進捗率", r"進捗"],
        "pv": [r"PV", r"計画値"],
        "ev": [r"EV", r"出来高"],
        "ac": [r"AC", r"実績コスト"]
    }

    def __init__(self, file_path, sheet_name='WBS_EVM'):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.config_path = os.path.join(os.path.dirname(file_path), "wbs_structure.json")
        self.config = None

    def infer_structure(self):
        """
        エクセルのヘッダーをスキャンし、構造を推論する。
        """
        # ヘッダー特定のために最初の数行を読み込む
        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
        if self.sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{self.sheet_name}' not found.")
        
        ws = wb[self.sheet_name]
        rows = list(ws.iter_rows(min_row=1, max_row=5, values_only=True))
        
        # 1. ヘッダー行の特定 (ID列がある行を探す)
        header_row_idx = -1
        for i, row in enumerate(rows):
            if any(re.search(p, str(cell)) for cell in row if cell for p in self.DEFAULT_ROLES["id"]):
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            raise ValueError("Could not find header row (ID column missing).")

        header_row = rows[header_row_idx]
        # フェーズ行（ヘッダーの上の行）があるか確認
        phase_row = rows[header_row_idx - 1] if header_row_idx > 0 else [None] * len(header_row)

        config = {
            "sheet_name": self.sheet_name,
            "header_row": header_row_idx + 1,
            "data_start_row": header_row_idx + 2,
            "columns": {
                "common": {},
                "phases": []
            }
        }

        # 2. カラムの分類
        current_phase = None
        phase_data = None

        for i, (col_name, p_name) in enumerate(zip(header_row, phase_row)):
            if col_name is None: continue
            col_name = str(col_name)

            # フェーズの切り替わり検知
            if p_name and str(p_name).strip():
                if phase_data:
                    config["columns"]["phases"].append(phase_data)
                current_phase = str(p_name).strip()
                phase_data = {"name": current_phase, "mapping": {}}

            # 役割の推論
            matched_role = None
            for role, patterns in self.DEFAULT_ROLES.items():
                if any(re.search(p, col_name) for p in patterns):
                    matched_role = role
                    break
            
            if matched_role:
                if matched_role in ["id", "name"]:
                    config["columns"]["common"][matched_role] = col_name
                elif phase_data:
                    # 同じ役割が複数回出てくる場合は、出現順にマッピングされる
                    if matched_role not in phase_data["mapping"]:
                        phase_data["mapping"][matched_role] = col_name

        if phase_data:
            config["columns"]["phases"].append(phase_data)

        self.config = config
        return config

    def get_column_index(self, role, phase_idx=0):
        """
        役割とフェーズインデックスから列インデックス(0-based)を取得する。
        """
        if not self.config:
            self.infer_structure()
            
        # TODO: 実際の実装では JSON から読み込むロジックが必要
        return 0 # プレースホルダー
