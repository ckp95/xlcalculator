#%%
from contextlib import contextmanager
import xlwings as xw
from hypothesis import given, settings
from hypothesis.strategies import integers, floats, one_of, none

from tests.testing import assert_equivalent, Case, parametrize_cases
from tests.conftest import formula_env
from xlcalculator.xlfunctions.math import DEC2BIN, DEC2OCT


MAX_EXAMPLES = 10000


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

def near(value, width=5):
    return xl_numbers(min_value=value - width, max_value=value + width)
    

env_dec2bin = formula_env(DEC2BIN, "number")
    

def test_dec2bin(env_dec2bin):
    variables = given(number=xl_numbers(-550, 550))
    fuzz_scalars(env=env_dec2bin, variables=variables)


env_dec2bin_places = formula_env(DEC2BIN, ["number", "places"])


def test_dec2bin_with_places(env_dec2bin_places):
    variables = given(number=xl_numbers(-550, 550), places=xl_numbers(-5, 15))
    fuzz_scalars(env=env_dec2bin_places, variables=variables)



env_dec2oct = formula_env(DEC2OCT, "number")


def test_dec2oct(env_dec2oct):
    lower, upper = -2**29, 2**29
    variables = given(number=one_of(xl_numbers(), near(lower), near(upper)))
    fuzz_scalars(env=env_dec2oct, variables=variables)


env_dec2oct_places = formula_env(DEC2OCT, ["number", "places"])


def test_dec2oct_with_places(env_dec2oct_places):
    lower, upper = -2**29, 2**29
    variables = given(
        number=xl_numbers(lower-5, upper+5),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_dec2oct_places, variables=variables)