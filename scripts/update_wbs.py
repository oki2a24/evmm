import pandas as pd
import openpyxl
import re
import os
import tempfile
import shutil
from datetime import datetime
from scripts.check_wbs_integrity import WBSIntegrityChecker

def update_wbs_logic(file_path, func_id, phase, effort, date=None):
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
    """
    if date is None:
        date = datetime.now()

    # 1. カラム位置の特定のために一度 pandas で読み込む
    # テンプレート構造上、1行目はフェーズ名、2行目が具体的な項目名となっているため、skiprows=1 を指定
    df = pd.read_excel(file_path, sheet_name='WBS_EVM', skiprows=1)
    cols = df.columns.tolist()

    def get_col_idx(name, occurrence=0):
        """
        指定された項目名の N 番目の出現インデックスを返す。
        正規表現により、pandasの自動サフィックス（.1, .2等）を許容。
        """
        pattern = re.compile(rf"^{re.escape(name)}(\.\d+)?$")
        matches = [i for i, c in enumerate(cols) if pattern.match(str(c))]
        if len(matches) > occurrence:
            return matches[occurrence]
        return None

    # フェーズの自動特定ロジック
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
        
        # 依存関係に基づき、未完了の最初のフェーズを特定
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
    
    # Excel上での行番号: index + 3 (skiprows=1, header=2, Excelは1-based)
    excel_row = target_row_idx[0] + 3

    # --- トランザクション処理 (Atomic Update) ---
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(file_path))
    os.close(fd)
    
    try:
        shutil.copy2(file_path, temp_path)

        wb = openpyxl.load_workbook(temp_path)
        ws = wb['WBS_EVM']

        # 開始日が未入力なら今日の日付を設定
        if ws.cell(row=excel_row, column=idx_sa + 1).value is None:
            ws.cell(row=excel_row, column=idx_sa + 1, value=date)
        
        ws.cell(row=excel_row, column=idx_ea + 1, value=date)
        ws.cell(row=excel_row, column=idx_ma + 1, value=effort)
        ws.cell(row=excel_row, column=idx_pr + 1, value=100)

        wb.save(temp_path)
        wb.close()

        # 整合性チェック (Hard Gate)
        checker = WBSIntegrityChecker(temp_path)
        df_for_check = checker.load_wbs()
        errors = checker.check_dataframe(df_for_check)
        if errors:
            error_msg = "\n".join([f"行 {e['index']+3}: {e['message']}" for e in errors])
            raise ValueError(f"整合性チェックエラーが発生しました。更新を中断します（ファイルは保護されました）:\n{error_msg}")

        # 検証成功時のみ元のファイルを上書き
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
    parser.add_argument('--file', default='projects/test_project/hoge_wbs_evm.xlsx', help='Excelファイルパス')
    
    args = parser.parse_args()
    try:
        update_wbs_logic(args.file, args.id, args.phase, args.effort)
        print(f"SUCCESS: {args.id} を更新しました。")
    except Exception as e:
        print(f"ERROR: {e}")
        exit(1)
