import datetime
import pytest
import os
import shutil
from scripts.utils import get_working_days, ensure_project_context

def test_ensure_project_context_creates_dir_and_file(tmp_path):
    """
    プロジェクトパスを渡し、docs/context.md が生成されることを検証する。
    """
    project_path = tmp_path / "test_project"
    project_path.mkdir()
    
    context_path = project_path / "docs" / "context.md"
    
    # 実行前は存在しないことを確認
    assert not context_path.exists()
    
    # 実行
    ensure_project_context(str(project_path))
    
    # 実行後は存在し、初期内容が含まれていることを確認
    assert context_path.exists()
    content = context_path.read_text()
    assert "# Project Context" in content
    assert "## 1. PJ特性 & 重要事項" in content
    assert "## 2. チーム / リソース状況" in content
    assert "## 3. 各タスクの背景・洞察" in content

def test_get_working_days_no_holidays():
    """
    祝日を含まない期間の稼働日計算を検証する。
    2026-04-01 (水) to 2026-04-03 (金) -> 3 days
    """
    start = datetime.date(2026, 4, 1)
    end = datetime.date(2026, 4, 3)
    assert get_working_days(start, end) == 3

def test_get_working_days_with_weekend():
    """
    週末（土日）を跨ぐ期間の稼働日計算を検証する。
    2026-04-01 (水) to 2026-04-06 (月) -> 4 working days (Wed, Thu, Fri, Mon)
    """
    start = datetime.date(2026, 4, 1)
    end = datetime.date(2026, 4, 6)
    assert get_working_days(start, end) == 4

def test_get_working_days_with_holidays_gw():
    """
    日本の祝日（ゴールデンウィーク）を含む期間の稼働日計算を検証する。
    2026-05-01 (金) to 2026-05-07 (木)
    5/1(金) : 稼働
    5/2(土), 5/3(日), 5/4(月), 5/5(火), 5/6(水) : 休日/祝日
    5/7(木) : 稼働
    期待値: 2日
    """
    start = datetime.date(2026, 5, 1)
    end = datetime.date(2026, 5, 7)
    assert get_working_days(start, end) == 2

def test_get_working_days_invalid_range():
    """
    開始日が終了日より後の場合、稼働日が0になることを検証する。
    """
    start = datetime.date(2026, 4, 2)
    end = datetime.date(2026, 4, 1)
    assert get_working_days(start, end) == 0
