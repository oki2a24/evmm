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

    def generate(self):
        """Excelファイルを生成し保存する"""
        self._create_settings_sheet()
        
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self.wb.save(self.output_path)
        print(f"Template generated at: {self.output_path}")

if __name__ == "__main__":
    # デフォルトの出力パス
    OUTPUT_PATH = "templates/master_template.xlsx"
    generator = TemplateGenerator(OUTPUT_PATH)
    generator.generate()
