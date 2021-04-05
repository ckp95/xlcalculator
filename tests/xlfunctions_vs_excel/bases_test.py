#%%
from contextlib import contextmanager
import xlwings as xw
from hypothesis import given, settings
from hypothesis.strategies import integers, floats, one_of, none, text

from tests.testing import assert_equivalent, Case, parametrize_cases
from tests.conftest import formula_env
from xlcalculator.xlfunctions.math import DEC2BIN, DEC2OCT, DEC2HEX, BIN2OCT, BIN2DEC, BIN2HEX, OCT2BIN, OCT2DEC, OCT2HEX, HEX2BIN


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
        none(),  # blank cell
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



env_dec2hex = formula_env(DEC2HEX, "number")


def test_dec2hex(env_dec2hex):
    lower, upper = -2**39, 2**39
    variables = given(number=one_of(xl_numbers(), near(lower), near(upper)))
    fuzz_scalars(env=env_dec2hex, variables=variables)
    

env_dec2hex_places = formula_env(DEC2HEX, ["number", "places"])


def test_dec2hex_with_places(env_dec2hex_places):
    lower, upper = -2**39, 2**39
    variables = given(
        number=xl_numbers(lower-5, upper+5),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_dec2hex_places, variables=variables)
    

def binary_numbers_as_strings():
    return text(
        alphabet=set("01"),
        min_size=1,
        max_size=11
    ).filter(lambda x: x == "0" or x[0] != "0")


env_bin2oct = formula_env(BIN2OCT, "number")

def test_bin2oct(env_bin2oct):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings()) 
    )
    fuzz_scalars(env=env_bin2oct, variables=variables)


env_bin2oct_places = formula_env(BIN2OCT, ["number", "places"])


def test_bin2oct_with_places(env_bin2oct_places):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings()),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_bin2oct_places, variables=variables)
    

env_bin2dec = formula_env(BIN2DEC, "number")


def test_bin2dec(env_bin2dec):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings()) 
    )
    fuzz_scalars(env=env_bin2dec, variables=variables)
    

# ___2DEC functions do not take a `places` parameter


env_bin2hex = formula_env(BIN2HEX, "number")


def test_bin2hex(env_bin2hex):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings()) 
    )
    fuzz_scalars(env=env_bin2hex, variables=variables)
    

env_bin2hex_places = formula_env(BIN2HEX, ["number", "places"])


def test_bin2hex_with_places(env_bin2hex_places):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings()),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_bin2hex_places, variables=variables)


def octal_numbers_as_strings(max_size):
    return text(
        alphabet=set("01234567"),
        min_size=1,
        max_size=max_size
    ).filter(lambda x: x == "0" or x[0] != "0")


env_oct2bin = formula_env(OCT2BIN, "number")


def test_oct2bin(env_oct2bin):
    variables = given(
        number=one_of(xl_numbers(), octal_numbers_as_strings(4)) 
    )
    fuzz_scalars(env=env_oct2bin, variables=variables)
    

env_oct2bin_places = formula_env(OCT2BIN, ["number", "places"])


def test_oct2bin_with_places(env_oct2bin_places):
    variables = given(
        number=one_of(xl_numbers(), octal_numbers_as_strings(4)),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_oct2bin_places, variables=variables)
    

env_oct2dec = formula_env(OCT2DEC, "number")

def test_oct2dec(env_oct2dec):
    variables = given(
        number=one_of(
            xl_numbers(),
            octal_numbers_as_strings(11)
        )
    )
    fuzz_scalars(env=env_oct2dec, variables=variables)
    

env_oct2hex = formula_env(OCT2HEX, "number")

def test_oct2hex(env_oct2hex):
    variables = given(
        number=one_of(
            xl_numbers(),
            octal_numbers_as_strings(11)
        )
    )
    fuzz_scalars(env=env_oct2hex, variables=variables)
    
    
env_oct2hex_places = formula_env(OCT2HEX, ["number", "places"])

def test_oct2hex_with_places(env_oct2hex_places):
    variables = given(
        number=one_of(xl_numbers(), octal_numbers_as_strings(11)),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_oct2hex_places, variables=variables)


def hex_numbers_as_strings(max_size):
    return text(
        alphabet=set("0123456789ABCDEFabcdef"),
        min_size=1,
        max_size=max_size
    ).filter(lambda x: x == "0" or x[0] != "0")

env_hex2bin = formula_env(HEX2BIN, "number")

def test_hex2bin(env_hex2bin):
    variables = given(
        number=one_of(
            xl_numbers(),
            hex_numbers_as_strings(3)
        )
    )
    fuzz_scalars(env=env_hex2bin, variables=variables)
    

env_hex2bin_places = formula_env(HEX2BIN, ["number", "places"])

def test_hex2bin_with_places(env_hex2bin_places):
    variables = given(
        number=one_of(xl_numbers(), hex_numbers_as_strings(3)),
        places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_hex2bin_places, variables=variables)


# todo: put max parameter on the binary strings too