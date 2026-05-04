import pandas as pd
import openpyxl
from scripts.utils import get_working_days

class WBSIntegrityChecker:
    """
    WBSの論理的整合性をチェックするクラス。
    """
    def __init__(self, file_path):
        self.file_path = file_path
        self.errors = []

    def load_wbs(self):
        """ExcelからWBSデータを読み込む。ヘッダー2行目を想定。"""
        # WBS_EVMシートを読み込み、2行目(skiprows=1)をヘッダーとする
        df = pd.read_excel(self.file_path, sheet_name='WBS_EVM', skiprows=1)
        return df

    def check_dataframe(self, df):
        """DataFrameに対して整合性チェックを実行する。"""
        errors = []
        for index, row in df.iterrows():
            # NaNや空行をスキップ（機能IDが空の場合はスキップ）
            if pd.isna(row.get("機能ID")):
                continue

            # --- 基本的な日付チェック (作成フェーズ) ---
            start_plan = pd.Timestamp(row.get("開始日予定")) if not pd.isna(row.get("開始日予定")) else pd.NaT
            end_plan = pd.Timestamp(row.get("終了日予定")) if not pd.isna(row.get("終了日予定")) else pd.NaT

            if pd.isna(start_plan):
                errors.append({"index": index, "message": "開始日予定が未入力です。"})
            if pd.isna(end_plan):
                errors.append({"index": index, "message": "終了日予定が未入力です。"})

            if not pd.isna(start_plan) and not pd.isna(end_plan):
                if start_plan > end_plan:
                    errors.append({"index": index, "message": f"開始日予定が終了日予定より後になっています({start_plan.date()} > {end_plan.date()})。"})

        return errors

    def run(self):
        """実行メイン処理"""
        df = self.load_wbs()
        self.errors = self.check_dataframe(df)
        
        if not self.errors:
            print("整合性チェック完了: 問題は見つかりませんでした。")
        else:
            print(f"整合性チェック完了: {len(self.errors)} 件の問題が見つかりました。")
            for e in self.errors:
                print(f"行 {e['index'] + 3}: {e['message']}") # Excel行番号に合わせる(+3)

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python scripts/check_wbs_integrity.py <excel_file_path>")
        sys.exit(1)
    
    checker = WBSIntegrityChecker(sys.argv[1])
    checker.run()
