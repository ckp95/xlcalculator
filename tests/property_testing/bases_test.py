#%%
from contextlib import contextmanager
import xlwings as xw
from hypothesis import given, settings
from hypothesis.strategies import integers, floats, one_of, none

from tests.testing import assert_equivalent, Case, parametrize_cases
from tests.conftest import formula_env
from xlcalculator.xlfunctions.math import DEC2BIN


MAX_EXAMPLES = 1000


def fuzz_scalars(env, variables, _settings=None):
    def inner(**kwargs):
        python_result = env.formula(**kwargs)
        env.set_args(**kwargs)
        excel_result = env.value
        
        assert_equivalent(result=python_result, expected=excel_result)
    
    _settings = _settings or settings(deadline=None, max_examples=MAX_EXAMPLES)
    func = variables(_settings(inner))
    func()
    

def xl_numbers(min_value=None, max_value=None):
    return one_of(
        none(),
        integers(min_value=min_value, max_value=max_value),
        floats(min_value=min_value, max_value=max_value, allow_infinity=False, allow_nan=False)
    )
    

env_dec2bin = formula_env(DEC2BIN, "number")
    

def test_dec2bin(env_dec2bin):
    variables = given(number=xl_numbers(-550, 550))
    fuzz_scalars(env=env_dec2bin, variables=variables)


env_dec2bin_places = formula_env(DEC2BIN, ["number", "places"])


def test_dec2bin_with_places(env_dec2bin_places):
    variables = given(number=xl_numbers(-550, 550), places=xl_numbers(-5, 15))
    fuzz_scalars(env=env_dec2bin_places, variables=variables)
