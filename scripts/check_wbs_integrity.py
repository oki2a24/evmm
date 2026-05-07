import pandas as pd
import openpyxl
import re
from collections import Counter
from scripts.utils import get_working_days, FormulaIntegrityManager

class WBSIntegrityChecker:
    """
    WBSの論理的整合性をチェックし、必要に応じて数式を修復するクラス。
    """
    def __init__(self, file_path, fix=False):
        self.file_path = file_path
        self.fix = fix
        self.errors = []
        self.formula_manager = FormulaIntegrityManager()

    def load_wbs(self):
        """
        ExcelからWBSデータを読み込む。
        """
        df = pd.read_excel(self.file_path, sheet_name='WBS_EVM', skiprows=1)
        return df

    def check_dataframe(self, df):
        """
        DataFrameに対して整合性チェックを実行する。
        """
        errors = []
        cols = df.columns.tolist()
        
        def get_col(name, occurrence=0):
            pattern = re.compile(rf"^{re.escape(name)}(\.\d+)?$")
            matches = [i for i, c in enumerate(cols) if pattern.match(str(c))]
            if len(matches) > occurrence:
                return matches[occurrence]
            return None

        idx_id = get_col("機能ID")
        idx_s0 = get_col("開始日予定", 0)
        idx_e0 = get_col("終了日予定", 0)
        idx_m0 = get_col("工数予定", 0)
        
        idx_s1 = get_col("開始日予定", 1)
        idx_e1 = get_col("終了日予定", 1)
        
        idx_s2 = get_col("開始日予定", 2)
        
        idx_progress = get_col("進捗率(%)", 0)
        idx_ea0 = get_col("終了日実績", 0)

        for index, row in df.iterrows():
            if idx_id is not None and pd.isna(row.iloc[idx_id]):
                continue

            s_p0 = pd.to_datetime(row.iloc[idx_s0]) if idx_s0 is not None and not pd.isna(row.iloc[idx_s0]) else pd.NaT
            e_p0 = pd.to_datetime(row.iloc[idx_e0]) if idx_e0 is not None and not pd.isna(row.iloc[idx_e0]) else pd.NaT
            m_p0 = row.iloc[idx_m0] if idx_m0 is not None else None
            
            s_p1 = pd.to_datetime(row.iloc[idx_s1]) if idx_s1 is not None and not pd.isna(row.iloc[idx_s1]) else pd.NaT
            e_p1 = pd.to_datetime(row.iloc[idx_e1]) if idx_e1 is not None and not pd.isna(row.iloc[idx_e1]) else pd.NaT
            
            s_p2 = pd.to_datetime(row.iloc[idx_s2]) if idx_s2 is not None and not pd.isna(row.iloc[idx_s2]) else pd.NaT
            
            progress = row.iloc[idx_progress] if idx_progress is not None else None
            e_a0 = pd.to_datetime(row.iloc[idx_ea0]) if idx_ea0 is not None and not pd.isna(row.iloc[idx_ea0]) else pd.NaT

            if not pd.isna(s_p0) and not pd.isna(e_p0):
                if s_p0 > e_p0:
                    errors.append({"index": index, "message": f"開始日予定が終了日予定より後になっています({s_p0.date()} > {e_p0.date()})。"})
            
            if pd.isna(s_p0):
                errors.append({"index": index, "message": "開始日予定が未入力です。"})
            if pd.isna(e_p0):
                errors.append({"index": index, "message": "終了日予定が未入力です。"})

            if not pd.isna(s_p0) and not pd.isna(e_p0) and m_p0 is not None and not pd.isna(m_p0):
                w_days = get_working_days(s_p0.date(), e_p0.date())
                if m_p0 > w_days:
                    errors.append({"index": index, "message": f"工数予定({m_p0})が稼働日数({w_days})を超過しています(過負荷)。"})

            if not pd.isna(e_p0) and not pd.isna(s_p1):
                if e_p0 > s_p1:
                    errors.append({"index": index, "message": f"作成フェーズの終了日({e_p0.date()})より前にレビューフェーズが開始({s_p1.date()})されています。"})
            
            if not pd.isna(e_p1) and not pd.isna(s_p2):
                if e_p1 > s_p2:
                    errors.append({"index": index, "message": f"レビューフェーズの終了日({e_p1.date()})より前に修正フェーズが開始({s_p2.date()})されています。"})

            if progress == 100.0 and pd.isna(e_a0):
                errors.append({"index": index, "message": "進捗率100%ですが、終了日実績が未入力です。"})

        # --- 5. 数式の整合性チェック (Consistency-Driven) ---
        if idx_s0 is not None:
            formula_cols = []
            for occ in range(3):
                s_col_idx = get_col("開始日予定", occ)
                if s_col_idx is not None:
                    # PV, EV, AC の相対位置 (scripts/generate_master_template.py 参照)
                    # 予定(4), 実績(7), 進捗率(8), PV(9), EV(10), AC(11) -> s_col+7, +8, +9
                    formula_cols.extend([s_col_idx + 7, s_col_idx + 8, s_col_idx + 9])

            wb = openpyxl.load_workbook(self.file_path, data_only=False)
            ws = wb['WBS_EVM']
            
            for col_idx in formula_cols:
                patterns = []
                for row in range(3, ws.max_row + 1):
                    cell_val = ws.cell(row=row, column=col_idx+1).value
                    template = self.formula_manager.extract_template(cell_val, row)
                    if template:
                        patterns.append(template)
                
                if patterns:
                    counter = Counter(patterns)
                    most_common = counter.most_common(2)
                    winner_pattern, count = most_common[0]
                    
                    if not (len(most_common) > 1 and most_common[0][1] == most_common[1][1]):
                        for row in range(3, ws.max_row + 1):
                            cell_val = ws.cell(row=row, column=col_idx+1).value
                            current_template = self.formula_manager.extract_template(cell_val, row)
                            if current_template != winner_pattern:
                                col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                                errors.append({
                                    "index": row - 3, 
                                    "message": f"数式の不整合を検知しました (列 {col_letter})。期待値テンプレート: {winner_pattern}"
                                })
            wb.close()

        return errors

    def write_results_to_excel(self):
        """
        チェック結果をExcelの「整合性チェック結果」列に書き戻す。
        """
        wb = openpyxl.load_workbook(self.file_path)
        if 'WBS_EVM' not in wb.sheetnames:
            return

        ws = wb['WBS_EVM']
        header_row = 2
        last_col = ws.max_column
        target_col = None

        for c in range(1, last_col + 1):
            if ws.cell(row=header_row, column=c).value == "整合性チェック結果":
                target_col = c
                break

        if target_col is None:
            target_col = last_col + 1
            ws.cell(row=header_row, column=target_col).value = "整合性チェック結果"
            source_cell = ws.cell(row=header_row, column=target_col - 1)
            target_cell = ws.cell(row=header_row, column=target_col)
            target_cell.font = openpyxl.styles.Font(bold=True)
            target_cell.alignment = openpyxl.styles.Alignment(horizontal='center')
            target_cell.fill = openpyxl.styles.PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

        for r in range(3, ws.max_row + 1):
            ws.cell(row=r, column=target_col).value = None

        for e in self.errors:
            excel_row = e['index'] + 3
            current_val = ws.cell(row=excel_row, column=target_col).value
            new_msg = e['message']
            if current_val:
                new_msg = f"{current_val} | {new_msg}"
            ws.cell(row=excel_row, column=target_col).value = new_msg
            ws.cell(row=excel_row, column=target_col).font = openpyxl.styles.Font(color="FF0000")

        wb.save(self.file_path)
        print(f"Excelにチェック結果を書き込みました: {self.file_path}")

    def run(self):
        """
        実行メイン処理
        """
        df = self.load_wbs()
        self.errors = self.check_dataframe(df)

        if not self.errors:
            print("整合性チェック完了: 問題は見つかりませんでした。")
        else:
            print(f"整合性チェック完了: {len(self.errors)} 件の問題が見つかりました。")
            for e in self.errors:
                print(f"行 {e['index'] + 3}: {e['message']}")
            
            if not self.fix:
                print("\n[HINT] 数式の不整合が検知されました。'scripts/check_wbs.sh --fix' を実行することで、多数決パターンに基づき数式を自動修復できます。")

        if self.fix and self.errors:
            import shutil
            bak_path = self.file_path + ".bak"
            shutil.copy2(self.file_path, bak_path)
            print(f"バックアップを作成しました: {bak_path}")

            wb = openpyxl.load_workbook(self.file_path)
            ws = wb['WBS_EVM']
            
            # 再度カラム特定
            cols = df.columns.tolist()
            matches = [i for i, c in enumerate(cols) if "開始日予定" in str(c)]
            formula_cols = []
            for s_col in matches:
                formula_cols.extend([s_col + 7, s_col + 8, s_col + 9])
            
            repaired = self.formula_manager.repair_sheet(ws, columns=[c+1 for c in formula_cols], start_row=3, end_row=ws.max_row)
            
            if repaired > 0:
                wb.save(self.file_path)
                print(f"成功: {repaired} 個のセルを修復しました。")
                print("修復後の再検証を実行中...")
                new_df = self.load_wbs()
                self.errors = self.check_dataframe(new_df)
                if not self.errors:
                    print("再検証の結果、全ての不整合が解消されました。")
            else:
                print("修復可能な不整合は見つかりませんでした。")

        self.write_results_to_excel()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='WBS整合性チェック・修復スクリプト')
    parser.add_argument('file', help='Excelファイルパス')
    parser.add_argument('--fix', action='store_true', help='数式の不整合を自動修復する')
    
    args = parser.parse_args()
    checker = WBSIntegrityChecker(args.file, fix=args.fix)
    checker.run()
