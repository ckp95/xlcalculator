#%%
from contextlib import contextmanager
import xlwings as xw
from hypothesis import given, settings
from hypothesis.strategies import integers, floats, one_of, none, text, booleans

from tests.testing import assert_equivalent, Case, parametrize_cases
from tests.conftest import formula_env
from xlcalculator.xlfunctions.engineering import (
    DEC2BIN,
    DEC2OCT,
    DEC2HEX,
    BIN2OCT,
    BIN2DEC,
    BIN2HEX,
    OCT2BIN,
    OCT2DEC,
    OCT2HEX,
    HEX2BIN,
    HEX2OCT,
    HEX2DEC,
)


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
        booleans(),
        integers(min_value=min_value, max_value=max_value),
        floats(
            min_value=min_value,
            max_value=max_value,
            allow_infinity=False,
            allow_nan=False,
        ),
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
    lower, upper = -(2 ** 29), 2 ** 29
    variables = given(number=one_of(xl_numbers(), near(lower), near(upper)))
    fuzz_scalars(env=env_dec2oct, variables=variables)


env_dec2oct_places = formula_env(DEC2OCT, ["number", "places"])


def test_dec2oct_with_places(env_dec2oct_places):
    lower, upper = -(2 ** 29), 2 ** 29
    variables = given(
        number=xl_numbers(lower - 5, upper + 5), places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_dec2oct_places, variables=variables)


env_dec2hex = formula_env(DEC2HEX, "number")


def test_dec2hex(env_dec2hex):
    lower, upper = -(2 ** 39), 2 ** 39
    variables = given(number=one_of(xl_numbers(), near(lower), near(upper)))
    fuzz_scalars(env=env_dec2hex, variables=variables)


env_dec2hex_places = formula_env(DEC2HEX, ["number", "places"])


def test_dec2hex_with_places(env_dec2hex_places):
    lower, upper = -(2 ** 39), 2 ** 39
    variables = given(
        number=xl_numbers(lower - 5, upper + 5), places=xl_numbers(-5, 15)
    )
    fuzz_scalars(env=env_dec2hex_places, variables=variables)


def binary_numbers_as_strings(max_size):
    return text(alphabet=set("01"), min_size=1, max_size=max_size).filter(
        lambda x: x == "0" or x[0] != "0"
    )


env_bin2oct = formula_env(BIN2OCT, "number")


def test_bin2oct(env_bin2oct):
    variables = given(number=one_of(xl_numbers(), binary_numbers_as_strings(11)))
    fuzz_scalars(env=env_bin2oct, variables=variables)


env_bin2oct_places = formula_env(BIN2OCT, ["number", "places"])


def test_bin2oct_with_places(env_bin2oct_places):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings(11)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_bin2oct_places, variables=variables)


env_bin2dec = formula_env(BIN2DEC, "number")


def test_bin2dec(env_bin2dec):
    variables = given(number=one_of(xl_numbers(), binary_numbers_as_strings(11)))
    fuzz_scalars(env=env_bin2dec, variables=variables)


# ___2DEC functions do not take a `places` parameter


env_bin2hex = formula_env(BIN2HEX, "number")


def test_bin2hex(env_bin2hex):
    variables = given(number=one_of(xl_numbers(), binary_numbers_as_strings(11)))
    fuzz_scalars(env=env_bin2hex, variables=variables)


env_bin2hex_places = formula_env(BIN2HEX, ["number", "places"])


def test_bin2hex_with_places(env_bin2hex_places):
    variables = given(
        number=one_of(xl_numbers(), binary_numbers_as_strings(11)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_bin2hex_places, variables=variables)


def octal_numbers_as_strings(max_size):
    return text(alphabet=set("01234567"), min_size=1, max_size=max_size).filter(
        lambda x: x == "0" or x[0] != "0"
    )


env_oct2bin = formula_env(OCT2BIN, "number")


def test_oct2bin(env_oct2bin):
    variables = given(number=one_of(xl_numbers(), octal_numbers_as_strings(4)))
    fuzz_scalars(env=env_oct2bin, variables=variables)


env_oct2bin_places = formula_env(OCT2BIN, ["number", "places"])


def test_oct2bin_with_places(env_oct2bin_places):
    variables = given(
        number=one_of(xl_numbers(), octal_numbers_as_strings(4)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_oct2bin_places, variables=variables)


env_oct2dec = formula_env(OCT2DEC, "number")


def test_oct2dec(env_oct2dec):
    variables = given(number=one_of(xl_numbers(), octal_numbers_as_strings(11)))
    fuzz_scalars(env=env_oct2dec, variables=variables)


env_oct2hex = formula_env(OCT2HEX, "number")


def test_oct2hex(env_oct2hex):
    variables = given(number=one_of(xl_numbers(), octal_numbers_as_strings(11)))
    fuzz_scalars(env=env_oct2hex, variables=variables)


env_oct2hex_places = formula_env(OCT2HEX, ["number", "places"])


def test_oct2hex_with_places(env_oct2hex_places):
    variables = given(
        number=one_of(xl_numbers(), octal_numbers_as_strings(11)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_oct2hex_places, variables=variables)


def hex_numbers_as_strings(max_size):
    return text(
        alphabet=set("0123456789ABCDEFabcdef"), min_size=1, max_size=max_size
    ).filter(lambda x: x == "0" or x[0] != "0")


env_hex2bin = formula_env(HEX2BIN, "number")


def test_hex2bin(env_hex2bin):
    variables = given(number=one_of(xl_numbers(), hex_numbers_as_strings(3)))
    fuzz_scalars(env=env_hex2bin, variables=variables)


env_hex2bin_places = formula_env(HEX2BIN, ["number", "places"])


def test_hex2bin_with_places(env_hex2bin_places):
    variables = given(
        number=one_of(xl_numbers(), hex_numbers_as_strings(3)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_hex2bin_places, variables=variables)


env_hex2oct = formula_env(HEX2OCT, "number")


def test_hex2oct(env_hex2oct):
    variables = given(number=one_of(xl_numbers(), hex_numbers_as_strings(11)))
    fuzz_scalars(env=env_hex2oct, variables=variables)


env_hex2oct_places = formula_env(HEX2OCT, ["number", "places"])


def test_hex2oct_with_places(env_hex2oct_places):
    variables = given(
        number=one_of(xl_numbers(), hex_numbers_as_strings(11)),
        places=xl_numbers(-5, 15),
    )
    fuzz_scalars(env=env_hex2oct_places, variables=variables)


env_hex2dec = formula_env(HEX2DEC, "number")


def test_hex2dec(env_hex2dec):
    variables = given(number=one_of(xl_numbers(), hex_numbers_as_strings(11)))
    fuzz_scalars(env=env_hex2dec, variables=variables)


def to_bin(x):
    return bin(x)[2:]


def to_oct(x):
    return oct(x)[2:]


def to_hex(x):
    return hex(x)[2:]


def ten_chars_or_fewer(x):
    return len(x) <= 10


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2oct_and_oct2bin_are_inverses(binary_string):
    # Note that in Python, bin(1023) == "0b1111111111", but in Excel, DEC2BIN(1023) gives
    # a #NUM! error while BIN2DEC(1111111111) == -1. So the numbers in the @given
    # decorator are not the same as how Excel interprets the input. This test just says
    # that OCT2BIN and BIN2OCT are inverses of each other, it says nothing about what
    # actual values they produce or what they mean. The same is true of the other inverse
    # property tests later on in this file.
    assert binary_string == OCT2BIN(BIN2OCT(binary_string))


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2hex_and_hex2bin_are_inverses(binary_string):
    assert binary_string == HEX2BIN(BIN2HEX(binary_string))


@given(binary_string=integers(min_value=0).map(to_bin).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_bin2dec_and_dec2bin_are_inverses(binary_string):
    assert binary_string == DEC2BIN(BIN2DEC(binary_string))


@given(octal_string=integers(min_value=0).map(to_oct).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_oct2dec_and_dec2oct_are_inverses(octal_string):
    assert octal_string == DEC2OCT(OCT2DEC(octal_string))


@given(octal_string=integers(min_value=0).map(to_oct).filter(ten_chars_or_fewer))
@settings(max_examples=MAX_EXAMPLES)
def test_oct2hex_and_hex2oct_are_inverses(octal_string):
    assert octal_string == HEX2OCT(OCT2HEX(octal_string))


@given(decimal_integer=integers(min_value=-2**39, max_value=2**39 - 1))
@settings(max_examples=MAX_EXAMPLES)
def test_hex2dec_and_dec2hex_are_inverses(decimal_integer):
    assert decimal_integer == HEX2DEC(DEC2HEX(decimal_integer))
