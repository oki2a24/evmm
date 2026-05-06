import pytest
from datetime import date
from scripts.analyze_evm import EVMAnalyst

"""
EVM分析エンジンのテストコード

【テストの目的】
本テストは、EVMにおけるPV（Planned Value: 計画値）の計算が、
PMBOK理論に基づき、かつ「稼働日（土日祝日を除く）」を考慮して
正しく線形按分されることを物理的に証明します。

【設計上の注意】
- 稼働日計算には既存の scripts/utils.py のロジックを利用することを前提とします。
- 基準日（Status Date）がタスク期間外（開始前、終了後）の場合の境界値も検証します。
"""

def test_calculate_pv_linear():
    """
    PVの線形按分計算のテスト（正常系：期間中）
    
    【シナリオ】
    - 開始予定: 2026-04-01 (水)
    - 終了予定: 2026-04-14 (火) 
      -> 土日を除くと 10稼働日 (04/01-03, 04/06-10, 04/13-14)
    - 工数予定: 10.0 人日
    - 基準日: 2026-04-07 (火) 
      -> 04/01から数えて 5稼働日経過
    【期待値】
    - PV = (経過稼働日 / 総稼働日) * 工数予定 = (5 / 10) * 10.0 = 5.0
    """
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 4, 7))
    
    start_date = date(2026, 4, 1)
    end_date = date(2026, 4, 14)
    planned_effort = 10.0
    
    pv = analyst.calculate_pv(start_date, end_date, planned_effort)
    assert pv == 5.0

def test_calculate_pv_before_start():
    """基準日が開始日より前の場合、PVは 0.0 であるべき"""
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 3, 31))
    pv = analyst.calculate_pv(date(2026, 4, 1), date(2026, 4, 10), 10.0)
    assert pv == 0.0

def test_calculate_pv_after_end():
    """基準日が終了日より後の場合、PVは全額（工数予定）であるべき"""
    analyst = EVMAnalyst(file_path="dummy.xlsx", status_date=date(2026, 4, 15))
    pv = analyst.calculate_pv(date(2026, 4, 1), date(2026, 4, 10), 10.0)
    assert pv == 10.0

def test_calculate_ev():
    """EV = 工数予定 * 進捗率"""
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    ev = analyst.calculate_ev(10.0, 50.0)
    assert ev == 5.0
    
    ev_zero = analyst.calculate_ev(10.0, 0.0)
    assert ev_zero == 0.0

def test_calculate_ac():
    """AC = 工数実績の値そのもの"""
    analyst = EVMAnalyst(file_path="dummy.xlsx")
    ac = analyst.calculate_ac(8.5)
    assert ac == 8.5
    
    ac_none = analyst.calculate_ac(None)
    assert ac_none == 0.0
