import datetime
import jpholiday

def get_working_days(start_date, end_date):
    """
    土日および日本の祝日を除いた稼働日数を計算する。
    
    【設計意図】
    複雑なベクトル演算やライブラリの期間計算関数を使用せず、1日ずつループして
    jpholiday.is_holiday() で判定するシンプルな実装を採用しています。
    理由は以下の通りです：
    1. 正確性: 日本の祝日は振替休日や国民の休日など規則が複雑ですが、jpholiday を
       用いて1日ずつ判定することで、エッジケースでの計算ミスを確実に防げます。
    2. 保守性: 判定条件（例：特定の会社休日を追加するなど）の変更が容易です。
    3. パフォーマンス: WBSの期間（通常数ヶ月程度）であれば、ループによる
       オーバーヘッドは無視できるほど小さいため、読みやすさを優先しました。

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
