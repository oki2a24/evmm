from openpyxl import Workbook
import os

"""
WBS/EVM マスターテンプレート生成スクリプト

このスクリプトは、プロジェクト管理に使用する標準的な WBS/EVM Excel テンプレートを生成します。
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.workbook.defined_name import DefinedName
import os

class TemplateGenerator:
    """テンプレート生成を管理するクラス"""
    
    def __init__(self, output_path):
        self.output_path = output_path
        self.wb = Workbook()
        # デフォルトのシートを削除して新しく作成
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]

    def _create_settings_sheet(self):
        """Settingsシートを作成し、基本情報と名前定義を設定する"""
        ws = self.wb.create_sheet("Settings")
        
        # タイトルと見出しのスタイル
        header_font = Font(bold=True, size=12)
        
        # プロジェクト基本情報
        ws["A1"] = "プロジェクト基本情報"
        ws["A1"].font = header_font
        
        ws["A2"] = "プロジェクト開始日"
        ws["B2"] = "2026/05/01" # デフォルト値
        
        # メンバーリスト
        ws["A3"] = "メンバーリスト"
        ws["A3"].font = header_font
        
        ws["A4"] = "メンバー名"
        ws["B4"] = "役割"
        ws["C4"] = "チームリーダー"
        
        # サンプルメンバー
        members = [
            ("Aさん", "リーダー", "Aさん"),
            ("Bさん", "メンバー", "Aさん"),
            ("Cさん", "メンバー", "Aさん"),
        ]
        for i, (name, role, leader) in enumerate(members, start=5):
            ws.cell(row=i, column=1, value=name)
            ws.cell(row=i, column=2, value=role)
            ws.cell(row=i, column=3, value=leader)
            
        # 名前定義の追加（メンバーリストの範囲）
        # 例: MEMBER_LIST = Settings!$A$5:$A$100
        member_range = f"Settings!$A$5:$A$100"
        defn = DefinedName("MEMBER_LIST", attr_text=member_range)
        self.wb.defined_names.add(defn)

    def _create_wbs_evm_sheet(self):
        """WBS_EVMシートを作成し、3フェーズ構造を構築する"""
        ws = self.wb.create_sheet("WBS_EVM")
        
        # ヘッダースタイル
        header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        header_font = Font(bold=True)
        alignment = Alignment(horizontal="center", vertical="center")
        
        # 1行目: フェーズヘッダー (セル結合)
        phases = [
            ("作成", 6, 14),           # F列(6)からN列(14)
            ("レビュー実施", 15, 23),   # O列(15)からW列(23)
            ("レビュー後修正", 24, 32)  # X列(24)からAF列(32)
        ]
        
        for name, start_col, end_col in phases:
            ws.cell(row=1, column=start_col, value=name)
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            ws.cell(row=1, column=start_col).alignment = alignment
            ws.cell(row=1, column=start_col).font = header_font
            ws.cell(row=1, column=start_col).fill = header_fill

        # 2行目: カラムヘッダー
        base_cols = ["No", "機能ID", "機能名称", "担当メンバー", "チームリーダー"]
        for i, col_name in enumerate(base_cols, start=1):
            ws.cell(row=2, column=i, value=col_name)
            
        phase_cols = [
            "開始日予定", "終了日予定", "工数予定", "開始日実績", "終了日実績", "工数実績",
            "PV (計画値)", "EV (出来高)", "AC (実績コスト)"
        ]
        
        for i, phase in enumerate(phases):
            start_col = phase[1]
            for j, col_name in enumerate(phase_cols):
                cell = ws.cell(row=2, column=start_col + j, value=col_name)
                cell.font = header_font
                cell.alignment = alignment

    def generate(self):
        """Excelファイルを生成し保存する"""
        self._create_settings_sheet()
        self._create_wbs_evm_sheet()
        
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self.wb.save(self.output_path)
        print(f"Template generated at: {self.output_path}")

if __name__ == "__main__":
    # デフォルトの出力パス
    OUTPUT_PATH = "templates/master_template.xlsx"
    generator = TemplateGenerator(OUTPUT_PATH)
    generator.generate()
