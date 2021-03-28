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
    app = xlwings.App(add_book=False, visible=True)
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
    """Given an xlwings workbook, a function, and the names of the arguments of that
    function, initializes cells on the workbook so that the function can be tested against
    its Excel formula equivalent.

    For example,

    >>> env = FormulaTestingEnvironment(
    ...     wb=<some blank workbook>, formula=DEC2BIN, argnames=["number", "places"]
    ... )

    We can then set the "number" argument via the `set_args` method:

    >>> env.set_args(number=123, places=9)

    And get the formula result back from the live Excel workbook:

    >>> env.value
    '001111011'

    At the moment, it supports only single-cell, named arguments.
    """
    
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


def formula_env(formula: Callable, argnames: Union[str, Sequence[str]]):
    """Parametrized fixture that creates a FormulaTestingEnvironment from a function and
    argument names. The workbook is provided by the `excel_workbook` fixture.
    """

    if isinstance(argnames, str):
        argnames = [argnames]

    @pytest.fixture
    def env(excel_workbook):
        return FormulaTestingEnvironment(
            wb=excel_workbook, formula=formula, argnames=argnames
        )

    return env
