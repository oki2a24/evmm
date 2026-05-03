from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName
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
            ("Dさん", "リーダー", "Dさん"),
            ("Eさん", "メンバー", "Dさん"),
            ("Fさん", "メンバー", "Dさん"),
        ]
        for i, (name, role, leader) in enumerate(members, start=5):
            ws.cell(row=i, column=1, value=name)
            ws.cell(row=i, column=2, value=role)
            ws.cell(row=i, column=3, value=leader)
            
        # 名前定義の追加（メンバーリストの範囲）
        member_range = f"Settings!$A$5:$A$100"
        defn = DefinedName("MEMBER_LIST", attr_text=member_range)
        self.wb.defined_names.add(defn)

    def _create_wbs_evm_sheet(self):
        """WBS_EVMシートを作成し、3フェーズ構造と計算式を構築する"""
        ws = self.wb.create_sheet("WBS_EVM")
        
        # ヘッダースタイル
        header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
        header_font = Font(bold=True)
        alignment = Alignment(horizontal="center", vertical="center")
        
        # 1行目: フェーズヘッダー (セル結合)
        # 1フェーズ10列構成に変更 (進捗率を追加)
        phases = [
            ("作成", 6, 15),           # F列(6)からO列(15)
            ("レビュー実施", 16, 25),   # P列(16)からY列(25)
            ("レビュー後修正", 26, 35)  # Z列(26)からAI列(35)
        ]
        
        for name, start_col, end_col in phases:
            ws.cell(row=1, column=start_col, value=name)
            ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            cell = ws.cell(row=1, column=start_col)
            cell.alignment = alignment
            cell.font = header_font
            cell.fill = header_fill

        # 2行目: カラムヘッダー
        base_cols = ["No", "機能ID", "機能名称", "担当メンバー", "チームリーダー"]
        for i, col_name in enumerate(base_cols, start=1):
            cell = ws.cell(row=2, column=i, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            
        phase_cols = [
            "開始日予定", "終了日予定", "工数予定", "開始日実績", "終了日実績", "工数実績",
            "進捗率(%)", "PV (計画値)", "EV (出来高)", "AC (実績コスト)"
        ]
        
        for i, phase in enumerate(phases):
            start_col = phase[1]
            for j, col_name in enumerate(phase_cols):
                cell = ws.cell(row=2, column=start_col + j, value=col_name)
                cell.font = header_font
                cell.alignment = alignment
                cell.fill = header_fill

        # 3行目から100行目まで数式を適用
        for row in range(3, 103):
            for i, phase in enumerate(phases):
                s_col = phase[1]
                
                # 相対参照用の列記号
                plan_start = ws.cell(row=row, column=s_col).column_letter
                plan_end = ws.cell(row=row, column=s_col+1).column_letter
                plan_cost = ws.cell(row=row, column=s_col+2).column_letter
                progress = ws.cell(row=row, column=s_col+6).column_letter
                actual_cost_input = ws.cell(row=row, column=s_col+5).column_letter
                
                # PV (計画値): 列 s_col + 7
                pv_formula = f'=IF(TODAY()<{plan_start}{row}, 0, IF(TODAY()>{plan_end}{row}, {plan_cost}{row}, {plan_cost}{row} * (TODAY()-{plan_start}{row})/({plan_end}{row}-{plan_start}{row}+1)))'
                ws.cell(row=row, column=s_col+7, value=pv_formula)
                
                # EV (出来高): 列 s_col + 8
                ev_formula = f'={plan_cost}{row} * {progress}{row}'
                ws.cell(row=row, column=s_col+8, value=ev_formula)
                
                # AC (実績コスト): 列 s_col + 9
                ac_formula = f'={actual_cost_input}{row}'
                ws.cell(row=row, column=s_col+9, value=ac_formula)

        # アラート（条件付き書式）の追加
        # PV列 (L, V, AF) にデモ用のアラート
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        for col in ["M", "W", "AG"]: # PV (計画値) の列 (1フェーズ10列構成での調整後)
            ws.conditional_formatting.add(f"{col}3:{col}103", CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

    def _create_team_evm_sheet(self):
        """チームEVMシートを作成し、リーダーごとの集計テーブルを構築する"""
        ws = self.wb.create_sheet("チームEVM")
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        
        # Settingsからリーダーリストを取得（簡易的に Aさん, Dさん を対象とする）
        leaders = ["Aさん", "Dさん"]
        
        current_row = 1
        for leader in leaders:
            # チーム名
            ws.cell(row=current_row, column=1, value=f"{leader}チーム").font = Font(bold=True, size=12)
            current_row += 1
            
            # 1. EVM指標集計テーブル
            ws.cell(row=current_row, column=1, value="EVM指標集計").font = header_font
            current_row += 1
            
            evm_headers = ["フェーズ", "PV (計画値)", "EV (出来高)", "AC (実績コスト)", "SV", "CV", "SPI", "CPI"]
            for i, h in enumerate(evm_headers, start=1):
                cell = ws.cell(row=current_row, column=i, value=h)
                cell.font = header_font
                cell.fill = header_fill
            current_row += 1
            
            phases = [("作成", "M", "N", "O"), ("レビュー実施", "W", "X", "Y"), ("レビュー後修正", "AG", "AH", "AI")]
            for phase_name, pv_col, ev_col, ac_col in phases:
                ws.cell(row=current_row, column=1, value=phase_name)
                # SUMIFS(WBS_EVM!PV列, WBS_EVM!リーダー列, リーダー名)
                ws.cell(row=current_row, column=2, value=f'=SUMIFS(WBS_EVM!{pv_col}:{pv_col}, WBS_EVM!$E:$E, "{leader}")')
                ws.cell(row=current_row, column=3, value=f'=SUMIFS(WBS_EVM!{ev_col}:{ev_col}, WBS_EVM!$E:$E, "{leader}")')
                ws.cell(row=current_row, column=4, value=f'=SUMIFS(WBS_EVM!{ac_col}:{ac_col}, WBS_EVM!$E:$E, "{leader}")')
                # SV = EV - PV, CV = EV - AC
                ws.cell(row=current_row, column=5, value=f'=C{current_row}-B{current_row}')
                ws.cell(row=current_row, column=6, value=f'=C{current_row}-D{current_row}')
                # SPI = EV / PV, CPI = EV / AC (エラー回避のため IFERROR 使用)
                ws.cell(row=current_row, column=7, value=f'=IFERROR(C{current_row}/B{current_row}, 1)')
                ws.cell(row=current_row, column=8, value=f'=IFERROR(C{current_row}/D{current_row}, 1)')
                current_row += 1
            
            current_row += 1 # スペース
            
            # 2. タスクメトリクステーブル
            ws.cell(row=current_row, column=1, value="タスクメトリクス").font = header_font
            current_row += 1
            
            met_headers = ["フェーズ", "総数", "仕掛かり(予定)", "仕掛かり(実績)", "完了(予定)", "完了(実績)"]
            for i, h in enumerate(met_headers, start=1):
                cell = ws.cell(row=current_row, column=i, value=h)
                cell.font = header_font
                cell.fill = header_fill
            current_row += 1
            
            # メトリクス用の数式 (COUNTIFS を使用)
            # 例: 完了(実績) = COUNTIFS(リーダー列, リーダー名, 終了日実績列, "<>")
            # 座標はフェーズごとに異なるため、ループ内で調整が必要だが、ここでは代表的な列を指定
            phase_info = [("作成", "F", "G", "I", "J"), ("レビュー実施", "P", "Q", "S", "T"), ("レビュー後修正", "Z", "AA", "AC", "AD")]
            for p_name, p_start, p_end, a_start, a_end in phase_info:
                ws.cell(row=current_row, column=1, value=p_name)
                # 総数: COUNTIFS(リーダー列, リーダー名, 予定工数列, ">0")
                ws.cell(row=current_row, column=2, value=f'=COUNTIFS(WBS_EVM!$E:$E, "{leader}", WBS_EVM!${p_start.replace("F", "H")}:${p_start.replace("F", "H")}, ">0")')
                # 完了(実績): COUNTIFS(リーダー列, リーダー名, 終了日実績列, "<>")
                ws.cell(row=current_row, column=6, value=f'=COUNTIFS(WBS_EVM!$E:$E, "{leader}", WBS_EVM!${a_end}:${a_end}, "<>")')
                current_row += 1

            current_row += 2 # 次のチームへのスペース

    def generate(self):
        """Excelファイルを生成し保存する"""
        self._create_settings_sheet()
        self._create_wbs_evm_sheet()
        self._create_team_evm_sheet()
        
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self.wb.save(self.output_path)
        print(f"Template generated at: {self.output_path}")

if __name__ == "__main__":
    # デフォルトの出力パス
    OUTPUT_PATH = "templates/master_template.xlsx"
    generator = TemplateGenerator(OUTPUT_PATH)
    generator.generate()
