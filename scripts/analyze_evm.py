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
