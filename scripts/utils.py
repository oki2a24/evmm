import datetime
import jpholiday

def get_working_days(start_date, end_date):
    """
    土日および日本の祝日を除いた稼働日数を計算する。
    
    Args:
        start_date (datetime.date): 開始日
        end_date (datetime.date): 終了日
        
    Returns:
        int: 稼働日数。開始日が終了日より後の場合は 0。
    """
    if start_date > end_date:
        return 0
        
    working_days = 0
    current_date = start_date
    while current_date <= end_date:
        # 土日(5, 6)でない、かつ祝日でない場合のみカウント
        if current_date.weekday() < 5 and not jpholiday.is_holiday(current_date):
            working_days += 1
        current_date += datetime.timedelta(days=1)
        
    return working_days
