import pandas as pd
import openpyxl
from scripts.utils import get_working_days

class WBSIntegrityChecker:
    """
    WBSの論理적整合性をチェックするクラス。
    """
    def __init__(self, file_path):
        self.file_path = file_path
        self.errors = []

    def load_wbs(self):
        """
        ExcelからWBSデータを読み込む。
        【背景】テンプレートの構造上、1行目はフェーズ名（作成、レビュー等）、
        2行目が具体的なカラム名（開始日予定等）となっているため、
        skiprows=1 を指定して実質的なヘッダーを読み込んでいます。
        """
        df = pd.read_excel(self.file_path, sheet_name='WBS_EVM', skiprows=1)
        return df

    def check_dataframe(self, df):
        """
        DataFrameに対して整合性チェックを実行する。
        
        【実装上の工夫と背景】
        1. カラム名の動的解決: 
           ExcelのWBS構造は同じ項目名（例：「開始日予定」）がフェーズごとに複数存在します。
           pandasで読み込むと、これらは「開始日予定」, 「開始日予定.1」, 「開始日予定.2」のように
           サフィックスが自動付与されます。これを正規表現で柔軟に捕捉し、出現順序(occurrence)で
           フェーズを特定するようにしています。
        2. 位置ベース(iloc)のアクセス:
           カラム名が将来的に微調整（例：スペースの有無）されても、正規表現マッチングでインデックスを
           一度取得してしまえば、以降は iloc により高速かつ確実にアクセスできます。
        3. 型の統一(pd.to_datetime):
           Excelから読み込まれた Timestamp 型と、手動作成されたデータや datetime.date 型が混在すると、
           比較演算( > )が正しく動作しない、あるいはエラーになる場合があります。
           これを防ぐため、日付比較の直前に全て pd.to_datetime で統一しています。
        """
        errors = []
        
        cols = df.columns.tolist()
        
        def get_col(name, occurrence=0):
            """
            指定された名前を含むカラムのインデックスを返す。
            pandasの自動サフィックス（.1, .2等）に対応するため正規表現を使用。
            """
            import re
            # 完全一致、または末尾に .数字 が付くものを探す (例: "開始日予定" や "開始日予定.1")
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
            # 機能IDが空の場合はスキップ
            if idx_id is not None and pd.isna(row.iloc[idx_id]):
                continue

            # --- 各フェーズの日付・工数を取得 ---
            s_p0 = pd.to_datetime(row.iloc[idx_s0]) if idx_s0 is not None and not pd.isna(row.iloc[idx_s0]) else pd.NaT
            e_p0 = pd.to_datetime(row.iloc[idx_e0]) if idx_e0 is not None and not pd.isna(row.iloc[idx_e0]) else pd.NaT
            m_p0 = row.iloc[idx_m0] if idx_m0 is not None else None
            
            s_p1 = pd.to_datetime(row.iloc[idx_s1]) if idx_s1 is not None and not pd.isna(row.iloc[idx_s1]) else pd.NaT
            e_p1 = pd.to_datetime(row.iloc[idx_e1]) if idx_e1 is not None and not pd.isna(row.iloc[idx_e1]) else pd.NaT
            
            s_p2 = pd.to_datetime(row.iloc[idx_s2]) if idx_s2 is not None and not pd.isna(row.iloc[idx_s2]) else pd.NaT
            
            progress = row.iloc[idx_progress] if idx_progress is not None else None
            e_a0 = pd.to_datetime(row.iloc[idx_ea0]) if idx_ea0 is not None and not pd.isna(row.iloc[idx_ea0]) else pd.NaT

            # --- 1. 日付逆転チェック ---
            if not pd.isna(s_p0) and not pd.isna(e_p0):
                if s_p0 > e_p0:
                    errors.append({"index": index, "message": f"開始日予定が終了日予定より後になっています({s_p0.date()} > {e_p0.date()})。"})
            
            if pd.isna(s_p0):
                errors.append({"index": index, "message": "開始日予定が未入力です。"})
            if pd.isna(e_p0):
                errors.append({"index": index, "message": "終了日予定が未入力です。"})

            # --- 2. 過負荷チェック ---
            if not pd.isna(s_p0) and not pd.isna(e_p0) and m_p0 is not None and not pd.isna(m_p0):
                w_days = get_working_days(s_p0.date(), e_p0.date())
                if m_p0 > w_days:
                    errors.append({"index": index, "message": f"工数予定({m_p0})が稼働日数({w_days})を超過しています(過負荷)。"})

            # --- 3. フェーズ間順序性チェック ---
            if not pd.isna(e_p0) and not pd.isna(s_p1):
                if e_p0 > s_p1:
                    errors.append({"index": index, "message": f"作成フェーズの終了日({e_p0.date()})より前にレビューフェーズが開始({s_p1.date()})されています。"})
            
            if not pd.isna(e_p1) and not pd.isna(s_p2):
                if e_p1 > s_p2:
                    errors.append({"index": index, "message": f"レビューフェーズの終了日({e_p1.date()})より前に修正フェーズが開始({s_p2.date()})されています。"})

            # --- 4. 実績整合性チェック ---
            if progress == 100.0 and pd.isna(e_a0):
                errors.append({"index": index, "message": "進捗率100%ですが、終了日実績が未入力です。"})

        return errors

    def run(self):
        """
        実行メイン処理
        【背景】表示する行番号に +3 している理由は以下の通りです：
        1. pandasのインデックスは 0 から開始される。
        2. Excelのヘッダーが 2行分存在する。
        3. Excelの行番号は 1 から開始される。
        したがって、0番目のデータ行は Excel上の 3行目に相当します。
        """
        df = self.load_wbs()
        self.errors = self.check_dataframe(df)
        
        if not self.errors:
            print("整合性チェック完了: 問題は見つかりませんでした。")
        else:
            print(f"整合性チェック完了: {len(self.errors)} 件の問題が見つかりました。")
            for e in self.errors:
                print(f"行 {e['index'] + 3}: {e['message']}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python scripts/check_wbs_integrity.py <excel_file_path>")
        sys.exit(1)
    
    checker = WBSIntegrityChecker(sys.argv[1])
    checker.run()
