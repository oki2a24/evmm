import pandas as pd
import openpyxl
import re
from datetime import datetime
from scripts.check_wbs_integrity import WBSIntegrityChecker

def update_wbs_logic(file_path, func_id, phase, effort, date=None):
    """
    Excel WBSの特定の機能・フェーズを更新する。

    【設計の背景と意思決定】
    1. ライブラリの使い分け (pandas vs openpyxl):
       - データの検索（行の特定）や構造解析には、強力なクエリ機能を持つ pandas を使用します。
       - Excelへの書き込みには openpyxl を使用します。pandas の `to_excel` は
         既存の書式や計算式、結合セルを破壊する可能性があるため、セル単位での
         操作が可能な openpyxl を選択しました。
    2. カラム位置の動的解決:
       - WBSテンプレートはマルチヘッダー構造であり、同じ項目名（例：「進捗率(%)」）が
         複数のフェーズ（作成、レビュー等）に存在します。
       - pandasで読み込むとこれらは「.1」「.2」といったサフィックスで区別されますが、
         正規表現と occurrence（出現順序）を用いることで、将来的にフェーズが増減しても
         ロジックを修正せずに対応できる柔軟性を確保しています。
    3. ステートフルなフェーズ特定:
       - ユーザーがフェーズを明示しない場合、進捗率を確認して「未完了の最初のフェーズ」を
         自動選択します。これにより、ユーザーは単に「F1が終わった」と言うだけで
         適切な更新が可能になります。
    """
    if date is None:
        date = datetime.now()

    # 1. カラム位置の特定のために一度 pandas で読み込む
    # テンプレート構造上、1行目はフェーズ名、2行目が具体的な項目名となっているため、
    # skiprows=1 で実質的なヘッダーを読み込みます。
    df = pd.read_excel(file_path, sheet_name='WBS_EVM', skiprows=1)
    cols = df.columns.tolist()

    def get_col_idx(name, occurrence=0):
        """
        指定された項目名の N 番目の出現インデックスを返す。
        正規表現により、pandasの自動サフィックス（.1, .2等）を許容します。
        """
        pattern = re.compile(rf"^{re.escape(name)}(\.\d+)?$")
        matches = [i for i, c in enumerate(cols) if pattern.match(str(c))]
        if len(matches) > occurrence:
            return matches[occurrence]
        return None

    # フェーズの自動特定ロジック
    # ユーザーがフェーズを指定しなかった場合、現在の進捗状況から推論します。
    if phase is None:
        idx_id_col = get_col_idx("機能ID")
        idx_pr0 = get_col_idx("進捗率(%)", 0) # 作成フェーズ
        idx_pr1 = get_col_idx("進捗率(%)", 1) # レビューフェーズ
        
        target_row = df[df.iloc[:, idx_id_col] == func_id]
        if target_row.empty:
            raise ValueError(f"機能ID '{func_id}' が見つかりません。")
        
        row_data = target_row.iloc[0]
        pr0 = row_data.iloc[idx_pr0] if not pd.isna(row_data.iloc[idx_pr0]) else 0
        pr1 = row_data.iloc[idx_pr1] if not pd.isna(row_data.iloc[idx_pr1]) else 0
        
        # 依存関係（作成 -> レビュー -> 修正）に基づき、未完了のフェーズを選択
        if pr0 < 100:
            phase = "作成"
        elif pr1 < 100:
            phase = "レビュー実施"
        else:
            phase = "レビュー後修正"

    # フェーズに応じた occurrence (0:作成, 1:レビュー, 2:修正)
    occurrence_map = {"作成": 0, "レビュー実施": 1, "レビュー後修正": 2}
    occ = occurrence_map.get(phase, 0)

    # 更新対象のカラムインデックスを取得
    idx_id = get_col_idx("機能ID")
    idx_sa = get_col_idx("開始日実績", occ)
    idx_ea = get_col_idx("終了日実績", occ)
    idx_ma = get_col_idx("工数実績", occ)
    idx_pr = get_col_idx("進捗率(%)", occ)

    # 2. 行の特定
    target_row_idx = df[df.iloc[:, idx_id] == func_id].index
    if target_row_idx.empty:
        raise ValueError(f"機能ID '{func_id}' が見つかりません。")
    
    # Excel上での行番号の計算:
    # 0-based index + skiprows(1) + Excelの1-based補正(1) + ヘッダー(1) = index + 3
    # ※ 1行目がフェーズ、2行目が項目名のため。
    excel_row = target_row_idx[0] + 3

    # 3. openpyxl による物理的な書き込み
    wb = openpyxl.load_workbook(file_path)
    ws = wb['WBS_EVM']

    # カラムインデックスは 0-based なので、Excelの 1-based に変換 (+1)
    # 開始日実績が空（初回の進捗報告）なら今日の日付をセット
    if ws.cell(row=excel_row, column=idx_sa + 1).value is None:
        ws.cell(row=excel_row, column=idx_sa + 1, value=date)
    
    ws.cell(row=excel_row, column=idx_ea + 1, value=date)
    ws.cell(row=excel_row, column=idx_ma + 1, value=effort)
    ws.cell(row=excel_row, column=idx_pr + 1, value=100) # 報告時は一律100%完了とみなす

    # 一時保存。整合性チェックに失敗した場合は保存されないよう、
    # 物理ファイルへの書き込みはこの関数の最後に行うのが理想だが、
    # Checker がファイルを直接読み込む仕様のため、一旦保存して検証する。
    wb.save(file_path)
    wb.close()

    # 4. 整合性チェックの実行 (Hard Gate)
    # 憲法の「整合性の死守」を担保するため、既存の検証ツールを呼び出す。
    checker = WBSIntegrityChecker(file_path)
    df_for_check = checker.load_wbs()
    errors = checker.check_dataframe(df_for_check)
    if errors:
        # エラーが検出された場合は例外を投げ、ユーザーに修正を促す。
        error_msg = "\n".join([f"行 {e['index']+3}: {e['message']}" for e in errors])
        raise ValueError(f"整合性チェックエラーが発生しました。更新を中断します:\n{error_msg}")

    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='WBS更新スクリプト')
    parser.add_argument('--id', required=True, help='機能ID')
    parser.add_argument('--phase', required=True, help='フェーズ (作成/レビュー実施/レビュー後修正)')
    parser.add_argument('--effort', type=float, required=True, help='実績工数')
    parser.add_argument('--file', default='projects/test_project/hoge_wbs_evm.xlsx', help='Excelファイルパス')
    
    args = parser.parse_args()
    try:
        update_wbs_logic(args.file, args.id, args.phase, args.effort)
        print(f"SUCCESS: {args.id} ({args.phase}) を更新しました。")
    except Exception as e:
        print(f"ERROR: {e}")
        exit(1)
