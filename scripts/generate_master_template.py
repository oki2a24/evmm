from openpyxl import Workbook
import os

"""
WBS/EVM マスターテンプレート生成スクリプト

このスクリプトは、プロジェクト管理に使用する標準的な WBS/EVM Excel テンプレートを生成します。
"""

class TemplateGenerator:
    """テンプレート生成を管理するクラス"""
    
    def __init__(self, output_path):
        self.output_path = output_path
        self.wb = Workbook()

    def generate(self):
        """Excelファイルを生成し保存する"""
        # 保存先ディレクトリが存在することを確認
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self.wb.save(self.output_path)
        print(f"Template generated at: {self.output_path}")

if __name__ == "__main__":
    # デフォルトの出力パス
    OUTPUT_PATH = "templates/master_template.xlsx"
    generator = TemplateGenerator(OUTPUT_PATH)
    generator.generate()
