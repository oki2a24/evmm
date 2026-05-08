import pandas as pd
import openpyxl
from datetime import date, datetime
from scripts.utils import get_working_days

"""
EVM分析エンジン (EVMAnalyst)

【クラスの役割】
本クラスは、WBSデータから正確なEVM指標（PV, EV, AC）を計算し、
プロジェクトの現状診断と将来予測を行うための「正確な計算機」として動作します。

【設計思想：正確な計算と将来の拡張性】
1. 稼働日ベースの計算: 単なる暦日ではなく、土日祝日を除外した「稼働日」を分母とすることで、
   PMが実感する進捗状況と一致したPV（計画値）を算出します。
2. 基準日（Status Date）の柔軟性: 分析の基準日を動的に指定可能にすることで、
   過去の特定時点での効率分析や、将来のシミュレーションを可能にします。
3. エクセルへの外科的更新: 既存のセルの書式や数式を壊さないよう、openpyxlを用いて
   計算結果のみを特定セルに書き込みます。
"""

class EVMAnalyst:
    def __init__(self, file_path, status_date=None):
        """
        EVM分析エンジンの初期化
        
        :param file_path: 分析対象のWBSエクセルファイルパス
        :param status_date: 分析基準日（デフォルトは今日）
        """
        self.file_path = file_path
        self.status_date = status_date if status_date else date.today()
        # 内部で扱う日付は比較のために datetime.date 型に統一します
        if isinstance(self.status_date, datetime):
            self.status_date = self.status_date.date()

    def calculate_pv(self, start_date, end_date, planned_effort):
        """
        PV（Planned Value：計画値）を線形按分法で算出する。
        
        【ロジックの背景】
        PVは「現時点で完了しているべき作業量」を示します。
        本実装では、タスクの総稼働日数に対する基準日までの経過稼働日数の割合を
        工数予定に乗じることで算出します。
        
        :param start_date: タスク開始予定日
        :param end_date: タスク終了予定日
        :param planned_effort: タスクの総工数予定（人日）
        :return: 基準日時点のPV
        """
        # 日付型の変換（Timestamp等への対応）
        if hasattr(start_date, 'date'): start_date = start_date.date()
        if hasattr(end_date, 'date'): end_date = end_date.date()
        
        # 1. 基準日が開始日前なら 0.0
        if self.status_date < start_date:
            return 0.0
        
        # 2. 基準日が終了日後なら 全額 (100%)
        if self.status_date >= end_date:
            return planned_effort
            
        # 3. 期間内の場合：稼働日ベースで按分
        # 総稼働日数の計算（開始から終了まで）
        total_working_days = get_working_days(start_date, end_date)
        if total_working_days <= 0:
            return 0.0
            
        # 経過稼働日数の計算（開始から基準日まで）
        elapsed_working_days = get_working_days(start_date, self.status_date)
        
        # 線形按分 (経過日数 / 総日数) * 総工数
        pv = (elapsed_working_days / total_working_days) * planned_effort
        return round(pv, 2)

    def calculate_ev(self, planned_effort, progress_rate):
        """
        EV（Earned Value：出来高）を算出する。
        
        :param planned_effort: 工数予定（人日）
        :param progress_rate: 進捗率（%）
        :return: 進捗に応じた出来高
        """
        if pd.isna(progress_rate):
            return 0.0
        return round(planned_effort * (progress_rate / 100.0), 2)

    def calculate_ac(self, actual_effort):
        """
        AC（Actual Cost：実績コスト）を算出する。
        本プロジェクトでは「工数実績」列に入力された人日をそのままACとして扱います。
        
        :param actual_effort: 工数実績（人日）
        """
        if pd.isna(actual_effort):
            return 0.0
        return float(actual_effort)

    def calculate_bac(self, df):
        """
        BAC（Budget At Completion：完成時総予算）を算出する。
        WBS全体の全フェーズの「工数予定」を合算します。
        
        :param df: WBSデータのDataFrame
        :return: 総予算（人日）
        """
        total_bac = 0.0
        # 「工数予定」で始まる全てのカラムを抽出して合算
        effort_cols = [c for c in df.columns if str(c).startswith("工数予定")]
        for col in effort_cols:
            total_bac += df[col].sum()
        return round(total_bac, 2)

    def calculate_forecasts(self, bac, ev, ac, pv):
        """
        PMBOK公式に基づき、3つの予測シナリオを算出する。
        
        :param bac: 完成時総予算
        :param ev: 出来高
        :param ac: 実績コスト
        :param pv: 計画値
        :return: 3シナリオの辞書
        """
        cpi = ev / ac if ac > 0 else 1.0
        spi = ev / pv if pv > 0 else 1.0
        
        # ゼロ除算防止のためのガードレール (最低効率を 0.01 に設定)
        safe_cpi = max(cpi, 0.01)
        safe_spi = max(spi, 0.01)
        
        # 1. 現実的 (Realistic): 現在の効率が継続
        eac_r = bac / safe_cpi
        
        # 2. 楽観的 (Optimistic): 残りは計画通り
        eac_o = ac + (bac - ev)
        
        # 3. 慎重 (Pessimistic): 現在の効率とスケジュールの両方を考慮
        eac_p = ac + ((bac - ev) / (safe_cpi * safe_spi))
        
        def build_res(eac, desc):
            return {
                "eac": round(eac, 2),
                "etc": round(max(eac - ac, 0), 2),
                "vac": round(bac - eac, 2),
                "description": desc
            }
            
        return {
            "realistic": build_res(eac_r, "現在の効率が継続した場合の予測"),
            "optimistic": build_res(eac_o, "残作業を計画通りに遂行した場合"),
            "pessimistic": build_res(eac_p, "現在の効率と遅延が相互に影響した場合")
        }

    def run(self):
        """
        分析のメイン実行フロー。
        
        【規律：分析前の整合性チェック】
        プロジェクト憲法に基づき、計算を行う前に必ず WBSIntegrityChecker を実行します。
        不整合なデータ（例：開始日が終了日より後）がある場合、誤ったCPI/SPIを
        出力することを防ぐため、処理を中断します。
        """
        from scripts.check_wbs_integrity import WBSIntegrityChecker
        
        print(f"--- 整合性チェックを開始します: {self.file_path} ---")
        checker = WBSIntegrityChecker(self.file_path)
        df = checker.load_wbs()
        errors = checker.check_dataframe(df)
        
        if errors:
            # エラーがある場合はExcelにエラー内容を書き込んで中断
            checker.errors = errors
            checker.write_results_to_excel()
            raise ValueError(f"整合性チェックでエラーが検出されました ({len(errors)}件)。修正してから再度実行してください。")
        
        print("整合性チェックOK。分析を開始します...")

        results = []
        # 各フェーズ（0:作成, 1:レビュー, 2:修正）のデータを特定して計算
        for index, row in df.iterrows():
            # 機能IDが空の行はデータ行ではないとみなしてスキップ
            if pd.isna(row.get("機能ID")):
                continue

            # 最大3フェーズ（テンプレートの標準構造）をループ処理
            for phase_idx in range(3):
                suffix = f".{phase_idx}" if phase_idx > 0 else ""
                
                start_col = f"開始日予定{suffix}"
                end_col = f"終了日予定{suffix}"
                effort_col = f"工数予定{suffix}"
                progress_col = f"進捗率(%){suffix}"
                actual_col = f"工数実績{suffix}"
                pv_col_name = f"PV (計画値){suffix}"

                # テンプレート構造が拡張され、カラムが存在しない場合はスキップ
                if start_col not in df.columns:
                    continue

                # 予定が未入力の場合はPV計算ができないためスキップ
                if pd.isna(row[start_col]) or pd.isna(row[end_col]):
                    continue

                pv = self.calculate_pv(row[start_col], row[end_col], row[effort_col])
                ev = self.calculate_ev(row[effort_col], row[progress_col])
                ac = self.calculate_ac(row[actual_col])

                results.append({
                    "row": index + 3, # Excel行番号（Header 2行 + 1-index = +3）
                    "pv": pv,
                    "ev": ev,
                    "ac": ac,
                    "excel_pv": row[pv_col_name], # エクセル上の現在値（比較用）
                    "phase_idx": phase_idx
                })

        # 3. Excelへの書き戻し（外科的更新）
        # 【重要】エクセルの数式を保護するため、デフォルトでは上書きを無効化します。
        # 必要に応じて引数等で有効化できるよう設計上の余地を残します。
        # self.write_results_to_excel(results)
        
        # 4. AIアナリスト向けのサマリーJSON出力
        summary = self.get_summary_json(results)
        import json
        print("\n--- 分析集計結果 (JSON) ---")
        print(json.dumps(summary, indent=2, ensure_ascii=False))
        print("---------------------------\n")
        
        print(f"分析が正常に完了しました。エクセルの数式は保護されました。")
        return summary

    def write_results_to_excel(self, results):
        """
        計算結果（PV, EV, AC）をExcelの該当セルに書き込む。
        """
        wb = openpyxl.load_workbook(self.file_path)
        if 'WBS_EVM' not in wb.sheetnames:
            raise ValueError("'WBS_EVM' シートが見つかりません。")
            
        ws = wb['WBS_EVM']
        header_row = 2
        
        def get_column_indices(target_name):
            indices = []
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=header_row, column=c).value
                if val == target_name:
                    indices.append(c)
            return indices

        pv_cols = get_column_indices("PV (計画値)")
        ev_cols = get_column_indices("EV (出来高)")
        ac_cols = get_column_indices("AC (実績コスト)")

        for res in results:
            row_idx = res["row"]
            phase_idx = res.get("phase_idx", 0)
            
            if phase_idx < len(pv_cols):
                ws.cell(row=row_idx, column=pv_cols[phase_idx]).value = res["pv"]
            if phase_idx < len(ev_cols):
                ws.cell(row=row_idx, column=ev_cols[phase_idx]).value = res["ev"]
            if phase_idx < len(ac_cols):
                ws.cell(row=row_idx, column=ac_cols[phase_idx]).value = res["ac"]

        wb.save(self.file_path)
        print(f"Excelへの計算結果反映が完了しました: {self.file_path}")

    def get_summary_json(self, results):
        """
        全体の集計を行い、AIアナリスト向けのJSONデータを生成する。
        """
        total_pv = sum(r["pv"] for r in results)
        total_ev = sum(r["ev"] for r in results)
        total_ac = sum(r["ac"] for r in results)
        
        # 乖離（Gap）の分析：エクセル値と計算値の差を特定
        gaps = []
        for r in results:
            # 数値として比較可能な場合のみ
            try:
                e_pv = float(r["excel_pv"]) if not pd.isna(r["excel_pv"]) else 0.0
                diff = abs(r["pv"] - e_pv)
                if diff > 0.01: # わずかな誤差を除外
                    gaps.append({"row": r["row"], "excel_pv": e_pv, "python_pv": r["pv"], "diff": diff})
            except (ValueError, TypeError):
                # 数式が入っている場合は float() 変換に失敗するが、それは「正常」
                continue

        cpi = total_ev / total_ac if total_ac > 0 else 1.0
        spi = total_ev / total_pv if total_pv > 0 else 1.0
        
        summary = {
            "status_date": str(self.status_date),
            "metrics": {
                "total_pv": round(total_pv, 2),
                "total_ev": round(total_ev, 2),
                "total_ac": round(total_ac, 2),
                "cpi": round(cpi, 2),
                "spi": round(spi, 2)
            },
            "gap_analysis": {
                "count": len(gaps),
                "details": gaps[:5] # 代表的な乖離のみ提示
            },
            "alerts": []
        }
        
        if cpi < 0.9: summary["alerts"].append("COST_EFFICIENCY_LOW")
        if spi < 0.9: summary["alerts"].append("SCHEDULE_DELAYED")
        if len(gaps) > 0: summary["alerts"].append("EXCEL_FORMULA_GAP_DETECTED")
        
        return summary

if __name__ == "__main__":
    import sys
    import argparse
    from datetime import datetime

    parser = argparse.ArgumentParser(description='WBS/EVM 自動分析エンジン')
    parser.add_argument('file_path', help='分析対象のエクセルファイルパス')
    parser.add_argument('--date', help='分析基準日 (YYYY-MM-DD)。指定がない場合は今日。')
    
    args = parser.parse_args()
    
    status_date_val = None
    if args.date:
        try:
            status_date_val = datetime.strptime(args.date, '%Y-%m-%d').date()
        except ValueError:
            print("エラー: 日付形式は YYYY-MM-DD で指定してください。")
            sys.exit(1)
    
    analyst_obj = EVMAnalyst(args.file_path, status_date=status_date_val)
    analyst_obj.run()
