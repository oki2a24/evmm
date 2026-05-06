import pandas as pd
import openpyxl
import re
import os
import tempfile
import shutil
from datetime import datetime
from scripts.check_wbs_integrity import WBSIntegrityChecker

def update_wbs_logic(file_path, func_id, phase, effort, date=None, progress=100):
    """
    Excel WBSの特定の機能・フェーズを更新する。

    【設計の背景と意思決定 (Architecture & Decisions)】
    1. ライブラリの使い分け (pandas vs openpyxl):
       - データの検索（行の特定）や構造解析には、強力なクエリ機能を持つ pandas を使用。
       - Excelへの書き込みには openpyxl を使用。pandas の `to_excel` は
         既存の書式や計算式、結合セルを破壊する可能性があるため、セル単位での
         操作が可能な openpyxl を選択し、既存レイアウトを「外科的」に維持しています。
    2. カラム位置の動的解決:
       - WBSテンプレートはマルチヘッダー構造であり、同じ項目名（例：「進捗率(%)」）が
         複数のフェーズ（作成、レビュー等）に存在します。
       - pandasで読み込むとこれらは「.1」「.2」といったサフィックスで区別されますが、
         正規表現と occurrence（出現順序）を用いることで、将来的にフェーズが増減しても
         ロジックを修正せずに対応できる柔軟性を確保しています。
    3. ステートフルなフェーズ特定:
       - ユーザーがフェーズを明示しない場合、進捗率を確認して「未完了の最初のフェーズ」を
         自動選択します。これにより、ユーザーは単に「F1が終わった」と言うだけで
         適切な更新が可能になり、UXを大幅に向上させています。
    4. 原子性の担保 (Atomic Integrity):
       - コードレビューに基づき、「一時ファイル検証」フローを導入。
       - 一時ファイルに対して更新・検証（整合性チェック）を行い、成功した場合のみ
         元のファイルを上書き。これにより、エラー時にファイルが不整合な状態で
         保存されることを物理的に防ぎ、憲法が定める「システムの整合性死守」を実現しています。
    5. 書式の維持 (Formatting Preservation):
       - 日付を書き込む際、openpyxl はデフォルトで標準日時形式を適用してしまいます。
       - 既存の表示形式（m月d日等）を維持するため、同一行の「予定日」セルから 
         number_format を継承する処理を導入しています。
    """
    if date is None:
        date = datetime.now()

    # 1. カラム位置の特定のために一度 pandas で読み込む
    df = pd.read_excel(file_path, sheet_name='WBS_EVM', skiprows=1)
    cols = df.columns.tolist()

    def get_col_idx(name, occurrence=0):
        pattern = re.compile(rf"^{re.escape(name)}(\.\d+)?$")
        matches = [i for i, c in enumerate(cols) if pattern.match(str(c))]
        if len(matches) > occurrence:
            return matches[occurrence]
        return None

    # フェーズの自動特定
    if phase is None:
        idx_id_col = get_col_idx("機能ID")
        idx_pr0 = get_col_idx("進捗率(%)", 0)
        idx_pr1 = get_col_idx("進捗率(%)", 1)
        
        target_row = df[df.iloc[:, idx_id_col] == func_id]
        if target_row.empty:
            raise ValueError(f"機能ID '{func_id}' が見つかりません。")
        
        row_data = target_row.iloc[0]
        pr0 = row_data.iloc[idx_pr0] if not pd.isna(row_data.iloc[idx_pr0]) else 0
        pr1 = row_data.iloc[idx_pr1] if not pd.isna(row_data.iloc[idx_pr1]) else 0
        
        if pr0 < 100:
            phase = "作成"
        elif pr1 < 100:
            phase = "レビュー実施"
        else:
            phase = "レビュー後修正"

    occurrence_map = {"作成": 0, "レビュー実施": 1, "レビュー後修正": 2}
    occ = occurrence_map.get(phase, 0)

    idx_id = get_col_idx("機能ID")
    idx_sp = get_col_idx("開始日予定", occ) # 継承元
    idx_sa = get_col_idx("開始日実績", occ)
    idx_ea = get_col_idx("終了日実績", occ)
    idx_ma = get_col_idx("工数実績", occ)
    idx_pr = get_col_idx("進捗率(%)", occ)

    target_row_idx = df[df.iloc[:, idx_id] == func_id].index
    if target_row_idx.empty:
        raise ValueError(f"機能ID '{func_id}' が見つかりません。")
    
    excel_row = target_row_idx[0] + 3

    # --- トランザクション処理 ---
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(file_path))
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)

        wb = openpyxl.load_workbook(temp_path)
        ws = wb['WBS_EVM']

        # 書式（表示形式）を予定日セルから継承する
        # ※ idx_sp (開始日予定) が見つからない場合はデフォルトの number_format を使用
        date_format = "m/d" # フォールバック
        if idx_sp is not None:
            date_format = ws.cell(row=excel_row, column=idx_sp + 1).number_format

        # 開始日実績の更新
        cell_sa = ws.cell(row=excel_row, column=idx_sa + 1)
        if cell_sa.value is None:
            cell_sa.value = date
            cell_sa.number_format = date_format
        
        # 終了日実績の更新
        cell_ea = ws.cell(row=excel_row, column=idx_ea + 1)
        cell_ea.value = date
        cell_ea.number_format = date_format
        
        ws.cell(row=excel_row, column=idx_ma + 1, value=effort)
        ws.cell(row=excel_row, column=idx_pr + 1, value=progress)

        wb.save(temp_path)
        wb.close()

        # 整合性チェック (Hard Gate)
        checker = WBSIntegrityChecker(temp_path)
        df_for_check = checker.load_wbs()
        errors = checker.check_dataframe(df_for_check)
        if errors:
            error_msg = "\n".join([f"行 {e['index']+3}: {e['message']}" for e in errors])
            raise ValueError(f"整合性チェックエラーが発生しました。更新を中断します（ファイルは保護されました）:\n{error_msg}")

        # 検証成功！上書き
        shutil.move(temp_path, file_path)
        return True

    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='WBS更新スクリプト')
    parser.add_argument('--id', required=True, help='機能ID')
    parser.add_argument('--phase', help='フェーズ (作成/レビュー実施/レビュー後修正)。省略時は進捗から自動判別。')
    parser.add_argument('--effort', type=float, required=True, help='実績工数')
    parser.add_argument('--date', help='実績終了日 (YYYY-MM-DD)。省略時は今日。')
    parser.add_argument('--progress', type=float, default=100, help='進捗率(%%)。デフォルトは100。')
    parser.add_argument('--file', default='projects/test_project/hoge_wbs_evm.xlsx', help='Excelファイルパス')
    
    args = parser.parse_args()
    try:
        update_date = datetime.strptime(args.date, '%Y-%m-%d') if args.date else None
        update_wbs_logic(args.file, args.id, args.phase, args.effort, date=update_date, progress=args.progress)
        print(f"SUCCESS: {args.id} を更新しました。")
    except Exception as e:
        print(f"ERROR: {e}")
        exit(1)
