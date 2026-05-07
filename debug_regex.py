import re

cell_ref_pattern = re.compile(r'(?<![!])(?P<col>[A-Za-z]{1,3})(?<!\$)(?P<row>\d+)')

def extract_template(formula, row):
    row_str = str(row)
    def repl(match):
        m_row = match.group('row')
        m_col = match.group('col').upper()
        if m_row == row_str:
            return f"{{{m_col}}}{{row}}"
        return match.group(0)
    return cell_ref_pattern.sub(repl, formula)

print(f"A3: {extract_template('=A3', 3)}")
print(f"VALUE3: {extract_template('=VALUE3', 3)}")
print(f"SUM1: {extract_template('=SUM1', 1)}")

cell_ref_pattern_fixed = re.compile(r'(?<![A-Za-z!])(?P<col>[A-Za-z]{1,3})(?<!\$)(?P<row>\d+)\b')

def extract_template_fixed(formula, row):
    row_str = str(row)
    def repl(match):
        m_row = match.group('row')
        m_col = match.group('col').upper()
        if m_row == row_str:
            return f"{{{m_col}}}{{row}}"
        return match.group(0)
    return cell_ref_pattern_fixed.sub(repl, formula)

print(f"FIXED A3: {extract_template_fixed('=A3', 3)}")
print(f"FIXED VALUE3: {extract_template_fixed('=VALUE3', 3)}")
print(f"FIXED SUM1: {extract_template_fixed('=SUM1', 1)}")
