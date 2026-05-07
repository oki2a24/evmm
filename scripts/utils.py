import datetime
import jpholiday
import re
from collections import Counter

def get_working_days(start_date, end_date):
    """
    土日および日本の祝日を除いた稼働日数を計算する。
    """
    if start_date > end_date:
        return 0
        
    working_days = 0
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5 and not jpholiday.is_holiday(current_date):
            working_days += 1
        current_date += datetime.timedelta(days=1)
        
    return working_days

class FormulaIntegrityManager:
    """
    Excelの数式整合性を管理し、パターンの抽出と修復を行うクラス。
    """
    def __init__(self):
        # セル参照（例: A3, $B$4, AA10）にマッチする正規表現
        # 1. (?<![A-Za-z!]): 直前に A-Z も ! もない（他シート参照および単語内マッチの防止）
        # 2. (?P<col>[A-Za-z]{1,3}): 列名は1〜3文字
        # 3. (?<!\$): 直前に $ がない（行の絶対参照を除外）
        # 4. (?P<row>\d+): 行番号
        # 5. (?![A-Za-z\d]): 直後に文字や数字がない（セル参照全体の境界を確保）
        self.cell_ref_pattern = re.compile(r'(?<![A-Za-z!])(?P<col>[A-Za-z]{1,3})(?<!\$)(?P<row>\d+)(?![A-Za-z\d])')

    def extract_template(self, formula, row):
        """
        数式から指定された行番号を {row} に置換してテンプレート化する。
        数式でない（=で始まらない）場合は None を返す。
        """
        if not isinstance(formula, str) or not formula.startswith('='):
            return None

        row_str = str(row)
        
        def repl(match):
            m_row = match.group('row')
            m_col = match.group('col').upper() # 大文字に正規化
            
            if m_row == row_str:
                return f"{{{m_col}}}{{row}}"
            return match.group(0)

        return self.cell_ref_pattern.sub(repl, formula)

    def repair_sheet(self, ws, columns, start_row, end_row, min_count=2):
        """
        指定された列をスキャンし、多数決で決まったパターンに基づいて数式を修復する。
        """
        total_repaired = 0
        
        for col_idx in columns:
            patterns = []
            # 1. パターンの収集
            for row in range(start_row, end_row + 1):
                cell_val = ws.cell(row=row, column=col_idx).value
                template = self.extract_template(cell_val, row)
                if template:
                    patterns.append(template)
            
            if not patterns:
                continue
                
            # 2. 多数決による正解パターンの決定
            counter = Counter(patterns)
            most_common = counter.most_common(2)
            
            winner_pattern, count = most_common[0]
            
            # 安全策: 出現数が閾値未満、またはタイの場合は修復しない
            if count < min_count:
                continue
            if len(most_common) > 1 and most_common[0][1] == most_common[1][1]:
                continue
            
            # 3. 修復の実行
            # カラム名の復元ロジック（{B} -> B）を事前に生成して最適化
            fixed_template = re.sub(r'\{([A-Z]+)\}', r'\1', winner_pattern)
            
            for row in range(start_row, end_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                current_template = self.extract_template(cell.value, row)
                
                if current_template != winner_pattern:
                    # 正解テンプレートに現在の行番号を流し込んで修復
                    new_formula = fixed_template.replace("{row}", str(row))
                    cell.value = new_formula
                    total_repaired += 1
                    
        return total_repaired
