# import pytest
# from _pytest.mark.structures import MarkDecorator
# from typing import Optional, Callable

# from xlcalculator.model import ModelCompiler
# from xlcalculator.evaluator import Evaluator
# from .. import testing
# from ..case import Case, parametrize_cases


from ..testing import workbook_test_cases, assert_equivalent





@workbook_test_cases("BASES.xlsx")
def test_bases(sheet_value, calculated_value):
    assert_equivalent(sheet_value, calculated_value, str)