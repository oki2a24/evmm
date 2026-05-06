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
        EV = 工数予定 * 進捗率(%)
        """
        if pd.isna(progress_rate):
            return 0.0
        return round(planned_effort * (progress_rate / 100.0), 2)

    def calculate_ac(self, actual_effort):
        """
        AC（Actual Cost：実績コスト）を算出する。
        本プロジェクトでは「工数実績」列の値をそのままACとして扱います。
        """
        if pd.isna(actual_effort):
            return 0.0
        return float(actual_effort)

    def run(self):
        """
        分析のメイン実行フロー。
        1. 整合性チェックの実行（必須）
        2. データのロードと計算
        3. Excelへの書き戻し
        """
        from scripts.check_wbs_integrity import WBSIntegrityChecker
        
        # 1. 整合性チェックの実行
        print(f"--- 整合性チェックを開始します: {self.file_path} ---")
        checker = WBSIntegrityChecker(self.file_path)
        df = checker.load_wbs()
        errors = checker.check_dataframe(df)
        
        if errors:
            # エラーがある場合はExcelに書き込んで中断
            checker.errors = errors
            checker.write_results_to_excel()
            raise ValueError(f"整合性チェックでエラーが検出されました ({len(errors)}件)。修正してから再度実行してください。")
        
        print("整合性チェックOK。分析を開始します...")

        # 2. データの計算
        results = []
        # 各フェーズ（0:作成, 1:レビュー, 2:修正）のデータを特定して計算
        # pandasの読み込みカラム名と一致させる
        for index, row in df.iterrows():
            # 機能IDが空の行はスキップ
            if pd.isna(row.get("機能ID")):
                continue

            # 各フェーズ（最大3フェーズを想定）について計算
            for phase_idx in range(3):
                suffix = f".{phase_idx}" if phase_idx > 0 else ""
                
                start_col = f"開始日予定{suffix}"
                end_col = f"終了日予定{suffix}"
                effort_col = f"工数予定{suffix}"
                progress_col = f"進捗率(%){suffix}"
                actual_col = f"工数実績{suffix}"

                # カラムが存在するかチェック
                if start_col not in df.columns:
                    continue

                pv = self.calculate_pv(row[start_col], row[end_col], row[effort_col])
                ev = self.calculate_ev(row[effort_col], row[progress_col])
                ac = self.calculate_ac(row[actual_col])

                results.append({
                    "row": index + 3, # Excel行番号（Header 2行 + 0-index 1行 = +3）
                    "pv": pv,
                    "ev": ev,
                    "ac": ac,
                    "phase_idx": phase_idx
                })

        # 3. Excelへの書き戻し
        self.write_results_to_excel(results)
        print(f"分析が正常に完了しました。")

    def write_results_to_excel(self, results):
        """
        計算結果（PV, EV, AC）をExcelの該当セルに書き込む。
        
        【外科的更新の重要性】
        既存のエクセルには複雑な数式や書式が含まれているため、pandas.to_excel() は使いません。
        openpyxlを用いて、計算した数値のみをピンポイントでセルに流し込みます。
        これにより、エクセルのダッシュボード機能（合計値の集計など）を活かしたまま、
        正確なデータを提供できます。
        
        :param results: リスト形式の計算結果。各要素は {"row": 行番号, "pv": 値, "ev": 値, "ac": 値, "phase_idx": 0/1/2}
        """
        wb = openpyxl.load_workbook(self.file_path)
        if 'WBS_EVM' not in wb.sheetnames:
            raise ValueError("'WBS_EVM' シートが見つかりません。")
            
        ws = wb['WBS_EVM']
        header_row = 2
        
        # ヘッダーから列インデックスを動的に取得するためのマッピング
        # カラム名のサフィックス（.1, .2）に対応するため、出現順序で管理します。
        def get_column_indices(target_name):
            indices = []
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=header_row, column=c).value
                # pandasが読み込む際のサフィックス形式（.1, .2等）ではなく、
                # エクセル上の生の文字列を比較します。
                # ただし、エクセル上では同じ名前が並んでいるだけなので、出現順にリスト化します。
                if val == target_name:
                    indices.append(c)
            return indices

        pv_cols = get_column_indices("PV (計画値)")
        ev_cols = get_column_indices("EV (出来高)")
        ac_cols = get_column_indices("AC (実績コスト)")

        # 結果を書き込み
        for res in results:
            row_idx = res["row"]
            phase_idx = res.get("phase_idx", 0) # デフォルトは最初のフェーズ
            
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
        AIアナリスト向けの集計データ（JSON）を生成する。
        """
        total_pv = sum(r["pv"] for r in results)
        total_ev = sum(r["ev"] for r in results)
        total_ac = sum(r["ac"] for r in results)
        
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
            "alerts": []
        }
        
        if cpi < 0.9: summary["alerts"].append("COST_EFFICIENCY_LOW")
        if spi < 0.9: summary["alerts"].append("SCHEDULE_DELAYED")
        
        return summary
