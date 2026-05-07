import pytest
from scripts.utils import FormulaIntegrityManager

def test_extract_formula_template_lowercase():
    """小文字の数式もテンプレート化できるべき"""
    manager = FormulaIntegrityManager()
    formula = "=b3+c3"
    # 現在の実装では [A-Z]+ なので失敗することが予想される
    template = manager.extract_template(formula, row=3)
    assert template == "={b}{row}+{c}{row}"

def test_extract_formula_template_sheet_reference():
    """シート参照が含まれる場合、シート名は抽象化されるべきではない"""
    manager = FormulaIntegrityManager()
    formula = "=Sheet1!A3"
    template = manager.extract_template(formula, row=3)
    # Sheet1 が {Sheet}{row} になると困る
    assert template == "=Sheet1!{A}{row}"

def test_extract_formula_template_multi_column():
    """2文字以上の列参照"""
    manager = FormulaIntegrityManager()
    formula = "=AA3+AB3"
    template = manager.extract_template(formula, row=3)
    assert template == "={AA}{row}+{AB}{row}"
