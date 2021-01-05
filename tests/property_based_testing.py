from dataclasses import dataclass
from typing import Sequence, Callable, Union
from types import SimpleNamespace

import xlwings
import pytest
from hypothesis import given, settings
from hypothesis.strategies import integers, floats

from xlcalculator.xlfunctions.math import DEC2BIN
from xlcalculator import xlerrors


@pytest.fixture(scope="session")
def excel_app():
    app = xlwings.App(add_book=False, visible=False)
    try:
        yield app
    finally:
        app.kill()


@pytest.fixture
def excel_workbook(excel_app):
    wb = excel_app.books.add()
    try:
        yield wb
    finally:
        wb.close()


error_lookup = {
    1: xlerrors.NullExcelError,
    2: xlerrors.DivZeroExcelError,
    3: xlerrors.ValueExcelError,
    4: xlerrors.RefExcelError,
    5: xlerrors.NameExcelError,
    6: xlerrors.NumExcelError,
    7: xlerrors.NaExcelError,
}


@dataclass
class FormulaTestingEnvironment:
    wb: xlwings.main.Book
    formula: Callable
    argnames: Sequence[str]
    
    def __post_init__(self):
        sheet = self.wb.sheets["Sheet1"]
        
        arg_addresses = {
            name: f"A{idx + 1}"
            for idx, name in enumerate(self.argnames)
        }
        self.args = {
            name: sheet[address]
            for name, address in arg_addresses.items()
        }
        
        args_string = ",".join(arg_addresses.values()).rstrip(",")
        formula_string = f"={self.formula_name}({args_string})"
        formula_address = "B1"
        self.formula_cell = sheet[formula_address]
        self.formula_cell.value = formula_string
        
        is_error_address = "C1"
        self.is_error_cell = sheet[is_error_address]
        self.is_error_cell.value = f"=ISERROR({formula_address})"
        
        error_type_address = "D1"
        self.error_type_cell = sheet[error_type_address]
        self.error_type_cell.value = f"=ERROR.TYPE({formula_address})"
        
    @property
    def formula_name(self):
        return self.formula.__name__
    
    @property
    def is_error(self):
        return self.is_error_cell.value
    
    @property
    def error_type(self):
        return error_lookup[self.error_type_cell.value]
    
    @property
    # awkward workaround because xlwings doesn't report the error code directly
    def value(self):
        if self.is_error:
            return self.error_type
        else:
            return self.formula_cell.value
    
    def set_args(self, **kwargs):        
        for name, value in kwargs.items():
            self.args[name].value = value

    
# def SQRT(x):
#     if x < 0:
#         return xlerrors.NumExcelError("cannot be less than 0")
    
#     return sqrt(x)


def formula_env(formula: Callable, argnames: Union[str, Sequence[str]]):
    if isinstance(argnames, str):
        argnames = [argnames]
    
    @pytest.fixture
    def env(excel_workbook):
        return FormulaTestingEnvironment(
            wb=excel_workbook, formula=formula, argnames=argnames
        )
    
    return env


# sqrt_env = formula_env(SQRT, "x")

# def test_sqrt_with_floats(sqrt_env):
#     @given(value=floats(allow_nan=False, allow_infinity=False))
#     @settings(max_examples=1000, deadline=None)
#     def inner(value):
#         result = SQRT(value)
#         sqrt_env.set_args(x=value)
#         expected = sqrt_env.value
        
#         assert_equivalent(result=result, expected=expected, normalize=isclose)
    
#     inner()
