import decimal
import math
from typing import Tuple, Union, Optional
from typing_extensions import Literal
from functools import wraps

import numpy as np
import pandas as pd
from scipy.special import factorial2

from . import xl, xlerrors, xlcriteria, func_xltypes
from .xlerrors import NumExcelError

# Testing Hook
rand = np.random.rand


@xl.register()
@xl.validate_args
def ABS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Find the absolute value of provided value.

    https://support.office.com/en-us/article/
        abs-function-3420200f-5628-4e8c-99da-c99d7c87713c
    """
    return abs(number)


@xl.register()
@xl.validate_args
def ACOS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arccosine, or inverse cosine, of a number.

    https://support.office.com/en-us/article/
        acos-function-cb73173f-d089-4582-afa1-76e5524b5d5b
    """
    return np.arccos(float(number))


@xl.register()
@xl.validate_args
def ACOSH(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the inverse hyperbolic cosine of a number.

    https://support.office.com/en-us/article/
        acosh-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe
    """
    if number < 1:
        raise xlerrors.NameExcelError(f'number {number} must be greater'
                                      f'than or equal to 1')

    return np.arccosh(float(number))


@xl.register()
@xl.validate_args
def ASIN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arcsine, or inverse sine, of a number.

    https://support.office.com/en-us/article/
        asin-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347
    """
    if number < -1 or number > 1:
        raise NumExcelError(f'number {number} must be less than '
                                     f'or equal to -1 or greater ot equal '
                                     f'to 1')

    return np.arcsin(float(number))


@xl.register()
@xl.validate_args
def ASINH(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the inverse hyperbolic sine of a number.

    https://support.office.com/en-us/article/
        asinh-function-4e00475a-067a-43cf-926a-765b0249717c
    """
    return np.arcsinh(float(number))


@xl.register()
@xl.validate_args
def ATAN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arctangent, or inverse tangent, of a number.

    https://support.office.com/en-us/article/
        atan-function-50746fa8-630a-406b-81d0-4a2aed395543
    """
    return np.arctan(float(number))


@xl.register()
@xl.validate_args
def ATAN2(
        x_num: func_xltypes.XlNumber,
        y_num: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the arctangent, or inverse tangent, of the specified
        x- and y-coordinates.

    https://support.office.com/en-us/article/
        atan2-function-c04592ab-b9e3-4908-b428-c96b3a565033
    """
    return np.arctan2(float(x_num), float(y_num))


@xl.register()
@xl.validate_args
def CEILING(
        number: func_xltypes.XlNumber,
        significance: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns number rounded up, away from zero, to the nearest multiple of
        significance.

    https://support.office.com/en-us/article/
        ceiling-function-0a5cd7c8-0720-4f0a-bd2c-c943e510899f
    """

    if significance == 0:
        return 0

    if significance < 0 < number:
        raise NumExcelError('significance below zero and number \
                                      above zero is not allowed')

    number = float(number)
    significance = float(significance)

    ceiling = significance * math.ceil(number / significance)

    # If number is an exact multiple of significance, no rounding occurs
    if (number % significance) == 0:
        return ceiling

    quantize_multiplier = str(significance % 1)

    # If number is negative, and significance is negative, the value is
    # rounded down, away from zero.
    if number < 0 and significance < 0:
        result = decimal.Decimal(ceiling)
        result = result.quantize(decimal.Decimal(quantize_multiplier),
                                 rounding=decimal.ROUND_DOWN)
        return float(result)

    # If number is negative, and significance is positive, the value is
    # rounded up towards zero.
    if number < 0 < significance:
        result = decimal.Decimal(ceiling)
        result = result.quantize(decimal.Decimal(quantize_multiplier),
                                 rounding=decimal.ROUND_UP)
        return float(result)

    # Regardless of the sign of number, a value is rounded up when adjusted
    # away from zero.
    result = decimal.Decimal(ceiling)
    result = result.quantize(decimal.Decimal(quantize_multiplier),
                             rounding=decimal.ROUND_UP)
    return float(result)


@xl.register()
@xl.validate_args
def COS(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the cosine of the given angle.

    https://support.office.com/en-us/article/
        cos-function-0fb808a5-95d6-4553-8148-22aebdce5f05
    """
    return np.cos(float(number))


@xl.register()
@xl.validate_args
def COSH(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the hyperbolic cosine of a number.

    https://support.office.com/en-us/article/
        cosh-function-e460d426-c471-43e8-9540-a57ff3b70555
    """
    return np.cosh(float(number))


@xl.register()
@xl.validate_args
def DEGREES(
        angle: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Converts radians into degrees.

    https://support.office.com/en-us/article/
        degrees-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1
    """
    return np.degrees(float(angle))


@xl.register()
@xl.validate_args
def EVEN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns number rounded up to the nearest even integer.

    https://support.office.com/en-us/article/
        even-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9

    algorithm found here;
        https://stackoverflow.com/questions/25361757/
            python-2-7-round-a-float-up-to-next-even-number
    """
    if number < 0:
        return math.ceil(abs(float(number)) / 2.) * -2
    else:
        return math.ceil(float(number) / 2.) * 2


@xl.register()
@xl.validate_args
def EXP(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns e raised to the power of number.

    https://support.office.com/en-us/article/
        exp-function-c578f034-2c45-4c37-bc8c-329660a63abe
    """
    return np.exp(float(number))


@xl.register()
@xl.validate_args
def FACT(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the factorial of a number

    https://support.office.com/en-us/article/
        fact-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3
    """
    if number < 0:
        raise NumExcelError('Negative values are not allowed')

    return math.factorial(int(number))


@xl.register()
@xl.validate_args
def FACTDOUBLE(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the double factorial of a number.

    https://support.office.com/en-us/article/
        factdouble-function-e67697ac-d214-48eb-b7b7-cce2589ecac8
    """
    if number < 0:
        raise NumExcelError('Negative values are not allowed')

    return factorial2(int(number), exact=True)


@xl.register()
@xl.validate_args
def FLOOR(
        number: func_xltypes.XlNumber,
        significance: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Rounds number down, toward zero, to the nearest multiple of
        significance.

    https://support.office.com/en-us/article/
        FLOOR-function-14BB497C-24F2-4E04-B327-B0B4DE5A8886
    """

    if significance < 0 < number:
        raise NumExcelError('number and significance needto have \
                                      the same sign')
    if number == 0:
        return 0

    if significance == 0:
        raise xlerrors.DivZeroExcelError()

    return significance * math.floor(number / significance)


@xl.register()
@xl.validate_args
def INT(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Rounds a number down to the nearest integer.

    https://support.office.com/en-us/article/
        int-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef
    """
    if number < 0:
        return _round(number, 0, _rounding=decimal.ROUND_UP)
    else:
        return _round(number, 0, _rounding=decimal.ROUND_DOWN)


@xl.register()
@xl.validate_args
def LN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the natural logarithm of a number.

    https://support.office.com/en-us/article/
        ln-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f
    """
    return math.log(number)


@xl.register()
@xl.validate_args
def LOG(
        number: func_xltypes.Number,
        base: func_xltypes.Number = 10
) -> func_xltypes.XlNumber:
    """Returns the logarithm of a number to the base you specify.

    https://support.office.com/en-us/article/
        log-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280
    """
    return math.log(float(number), float(base))


@xl.register()
@xl.validate_args
def LOG10(
        number: func_xltypes.Number
) -> func_xltypes.XlNumber:
    """Returns the base-10 logarithm of a number.

    https://support.office.com/en-us/article/
        log10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211
    """
    return np.log10(float(number))


@xl.register()
@xl.validate_args
def MOD(
        number: func_xltypes.XlNumber,
        divisor: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the remainder after number is divided by divisor.

    https://support.office.com/en-us/article/
        mod-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    """
    return number % divisor


@xl.register()
@xl.validate_args
def RAND() -> func_xltypes.XlNumber:
    """RAND returns an evenly distributed random real number greater than or
        equal to 0 and less than 1.

    https://support.office.com/en-us/article/
        rand-function-4cbfa695-8869-4788-8d90-021ea9f5be73
    """
    return rand()


@xl.register()
@xl.validate_args
def RANDBETWEEN(
        bottom: func_xltypes.XlNumber,
        top: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns a random integer number between the numbers you specify.

    https://support.office.com/en-us/article/
        randbetween-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685
    """
    return int(rand() * (top - bottom) + bottom)


@xl.register()
def PI() -> func_xltypes.XlNumber:
    """Returns the number 3.14159265358979, the mathematical constant pi.

    Accurate to 15 digits.

    https://support.office.com/en-us/article/
        pi-function-264199d0-a3ba-46b8-975a-c4a04608989b
    """
    return math.pi


@xl.register()
@xl.validate_args
def POWER(
        number: func_xltypes.XlNumber,
        power: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the result of a number raised to a power.

    https://support.office.com/en-us/article/
        power-function-d3f2908b-56f4-4c3f-895a-07fb519c362a
    """
    return np.power(number, power)


@xl.register()
@xl.validate_args
def RADIANS(
        angle: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Converts degrees to radians.

    https://support.office.com/en-us/article/
        radians-function-ac409508-3d48-45f5-ac02-1497c92de5bf
    """
    return np.radians(float(angle))


def _round(number, num_digits, _rounding=decimal.ROUND_HALF_UP):
    number = decimal.Decimal(str(number))
    dc = decimal.getcontext()
    dc.rounding = _rounding
    ans = round(number, int(num_digits))
    return float(ans)


@xl.register()
@xl.validate_args
def ROUND(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0,
        _rounding=decimal.ROUND_HALF_UP
):
    """Rounding half up

    https://support.office.com/en-us/article/
        ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c
    """
    return _round(number=number, num_digits=num_digits, _rounding=_rounding)


@xl.register()
@xl.validate_args
def ROUNDUP(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Round up

    https://support.office.com/en-us/article/
         ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7
    """
    return _round(number, num_digits=num_digits, _rounding=decimal.ROUND_UP)


@xl.register()
@xl.validate_args
def ROUNDDOWN(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Round down

    https://support.office.com/en-us/article/
        rounddown-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53
    """
    return _round(number, num_digits=num_digits, _rounding=decimal.ROUND_DOWN)


@xl.register()
@xl.validate_args
def SIGN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Determines the sign of a number.

    https://support.office.com/en-us/article/
        sign-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8
    """
    return np.sign(float(number))


@xl.register()
@xl.validate_args
def SIN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the sine of the given angle.

    https://support.office.com/en-us/article/
        sin-function-cf0e3432-8b9e-483c-bc55-a76651c95602
    """
    return float(np.sin(float(number)))


@xl.register()
@xl.validate_args
def SQRT(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns a positive square root.

    https://support.office.com/en-us/article/
        sqrt-function-654975c2-05c4-4831-9a24-2c65e4040fdf
    """
    if number < 0:
        raise NumExcelError(f'number {number} must be non-negative')

    return math.sqrt(number)


@xl.register()
@xl.validate_args
def SQRTPI(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the square root of (number * pi).

    https://support.office.com/en-us/article/
        sqrtpi-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4
    """
    if number < 0:
        raise NumExcelError(f'number {number} must be non-negative')

    return math.sqrt(number * math.pi)


@xl.register()
@xl.validate_args
def SUM(
        *numbers: Tuple[func_xltypes.XlNumber]
) -> func_xltypes.XlNumber:
    """The SUM function adds values.

    https://support.office.com/en-us/article/
        sum-function-043e1c7d-7726-4e80-8f32-07b23e057f89
    """
    # If no non numeric cells, return zero (is what excel does)
    if len(numbers) == 0:
        return 0

    return sum(numbers)


@xl.register()
@xl.validate_args
def SUMIF(
        range: func_xltypes.XlArray,
        criteria: func_xltypes.XlAnything,
        sum_range: func_xltypes.XlArray = None
) -> func_xltypes.XlNumber:
    """Adds the cells specified by a given criteria.

    https://support.office.com/en-us/article/
        sumif-function-169b8c99-c05c-4483-a712-1697a653039b
    """
    # WARNING:
    # - wildcards not supported

    check = xlcriteria.parse_criteria(criteria)

    if sum_range is None:
        sum_range = range

    range = range.flat
    sum_range = sum_range.cast_to_numbers().flat

    # zip() will automatically drop any range values that have indexes larger
    # than sum_range's length.
    return sum([
        sval
        for cval, sval in zip(range, sum_range)
        if check(cval)
    ])


@xl.register()
@xl.validate_args
def SUMIFS(
        sum_range: func_xltypes.XlArray,
        criteria_range: func_xltypes.XlArray,
        criteria: func_xltypes.XlAnything,
        *criteriaAndRanges: Tuple[Union[
            func_xltypes.XlAnything, func_xltypes.XlArray
        ]]
) -> func_xltypes.XlNumber:
    """Adds the cells specified by given criteria in multiple arrays.
    Requires equal length arrays, since the arrays get decomposed.

    https://support.microsoft.com/en-us/office
        /sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b
    """
    # WARNING:
    # - wildcards not supported
    ranges = [criteria_range.flat]
    checks = [xlcriteria.parse_criteria(criteria)]
    rangeLen = len(criteria_range.flat)
    newRange = []
    idx = 0
    for item in criteriaAndRanges:
        if idx == rangeLen:
            checks.append(xlcriteria.parse_criteria(item))
            ranges.append(newRange)
            newRange = []
            idx = 0
        else:
            newRange.append(item)
            idx += 1
    sum_range = sum_range.cast_to_numbers().flat
    # zip() will automatically drop any range values that have indexes larger
    # than sum_range's length.
    return sum([
        sval
        for cvals, sval in zip(zip(*ranges), sum_range)
        if all(checkfn(cvals[i]) for i, checkfn in enumerate(checks))
    ])


@xl.register()
@xl.validate_args
def SUMPRODUCT(
        *arrays: Tuple[func_xltypes.XlArray]
) -> func_xltypes.XlNumber:
    """Returns the sum of the products of corresponding arrays or arrays.

    https://support.office.com/en-us/article/
        sumproduct-function-16753e75-9f68-4874-94ac-4d2145a2fd2e
    """
    if len(arrays) == 0:
        raise xlerrors.NullExcelError('Not enough arguments for function.')

    array1_shape = arrays[0].shape
    if array1_shape == (0, 0):
        return 0

    for array in arrays:
        array_shape = array.shape
        if array1_shape != array_shape:
            raise xlerrors.ValueExcelError(
                f"The shapes of the arrays do not match. Looking "
                f"for {array1_shape} but given array has {array_shape}")
        if any(filter(xlerrors.ExcelError.is_error, xl.flatten(array))):
            raise xlerrors.NaExcelError(
                "Excel Errors are present in the sumproduct items.")

    sumproduct = pd.concat(arrays, axis=1)
    return sumproduct.prod(axis=1).sum()


@xl.register()
@xl.validate_args
def TAN(
        number: func_xltypes.XlNumber
) -> func_xltypes.XlNumber:
    """Returns the tangent of the given angle.

    https://support.office.com/en-us/article/
        tan-function-08851a40-179f-4052-b789-d7f699447401
    """
    return float(np.tan(float(number)))


@xl.register()
@xl.validate_args
def TRUNC(
        number: func_xltypes.XlNumber,
        num_digits: func_xltypes.XlNumber = 0
) -> func_xltypes.XlNumber:
    """Truncate a number to the specified number of digits.

    https://support.office.com/en-us/article/
        trunc-function-8b86a64c-3127-43db-ba14-aa5ceb292721
    """
    # Simple case. We want to make sure to return an integer in this
    # case.
    if num_digits == 0:
        return math.trunc(number)

    num_digits = int(num_digits)

    return math.trunc(number * 10**num_digits) / 10**num_digits


class Unused:
    """Some Excel formulae behave differently if you use a blank cell for one parameter vs
    not using that parameter at all. For example:

        =DEC2BIN(35)
        =DEC2BIN(35, A1)

    where A1 is a blank cell. The first will give "100011", the second gives a #NUM error.

    So, we have to be careful when using `None` as a default parameter value in the Python
    implementations. In cases where the distinction matters, use an instance of this
    `UNUSED` class as a default, instead of `None`.
    """
    pass


UNUSED = Unused()

def unused(x):
    return x is UNUSED


def used(x):
    return x is not UNUSED


@xl.register()
@xl.validate_args
def DEC2BIN(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if unused(places):
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    number = int(number)
    if not (-513 < number < 512):
        raise NumExcelError
    
    negative = number < 0
    bit_width = 10
    if negative:
        number += (1 << bit_width)
    
    string = bin(number)[2:]
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(desired_length)


@xl.register()
@xl.validate_args
def DEC2OCT(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if unused(places):
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    number = int(number)
    
    if not (-2**29 <= number < 2**29):
        raise NumExcelError
    
    bit_width = 30
    negative = number < 0
    
    if negative:
        number += (1 << bit_width)
        
    string = oct(number)[2:]
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(desired_length)


@xl.register()
@xl.validate_args
def DEC2HEX(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if unused(places):
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    number = int(number)
    
    if not (-2**39 <= number < 2**39):
        raise NumExcelError
    
    bit_width = 40
    negative = number < 0
    
    if negative:
        number += (1 << bit_width)
        
    string = hex(number)[2:].upper()
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)


@xl.register()
@xl.validate_args
def BIN2OCT(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if unused(places):
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    if not float(number).is_integer():
        raise NumExcelError
     
    number = int(number)
    if not (0 <= number < 10000000000): # could just check string length
        raise NumExcelError
    
    as_str = str(number)
    permitted_digits = set("01")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 2)
    
    bin_width = 10
    oct_width = 30
    
    mask = 1 << bin_width - 1
    
    # I figured this out months ago and now I don't remember why it works
    new_value = (value & ~mask) - (value & mask)
    
    negative = new_value < 0
    if negative:
        new_value += (1 << oct_width)
        
    # this doesn't have the `if new_value < 0 and new_value.bit_length() == 10 part`
    # because I can't come up with a counterexample where it is needed, but I imagine when
    # I come to refactor it can be put in the general case and not matter for BIN2OCT
    
    string = oct(new_value)[2:]
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)

 # the ___2DEC functions give a number not a string, and they do not take a `places`
 # parameter
@xl.register()
@xl.validate_args
def BIN2DEC(number: func_xltypes.XlNumber) -> func_xltypes.XlNumber:
    if not float(number).is_integer():
        raise NumExcelError
    
    number = int(number)
    if not (0 <= number < 10000000000):
        raise NumExcelError

    as_str = str(number)
    
    if set(as_str) - set("01"):
        raise NumExcelError
    
    bit_width = 10
    mask = 1 << bit_width - 1
    
    value = int(as_str, 2)
    return (value & ~mask) - (value & mask)


@xl.register()
@xl.validate_args
def BIN2HEX(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if unused(places):
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    
    if not float(number).is_integer():
        raise NumExcelError
    
    number = int(number)
    if not (0 <= number < 10000000000):
        raise NumExcelError
    
    as_str = str(number)
    permitted_digits = set("01")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 2)
    
    bin_width = 10
    hex_width = 40
    
    mask = 1 << bin_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    negative = new_value < 0
    if negative:
        new_value += (1 << hex_width)
    
    string = hex(new_value)[2:].upper()
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)


@xl.register()
@xl.validate_args
def OCT2BIN(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
        
    
    if not float(number).is_integer():
        raise NumExcelError
      
    number = int(number)
    
    as_str = str(number)
    permitted_digits = set("01234567")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 8)
    
    oct_width = 30
    bin_width = 10
    
    mask = 1 << oct_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    if not (-512 <= new_value < 512):
        raise NumExcelError
    
    negative = new_value < 0
    if negative:
        new_value += (1 << bin_width)
    
    string = bin(new_value)[2:]
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)
    
    
@xl.register()
@xl.validate_args
def OCT2DEC(number: func_xltypes.XlNumber) -> func_xltypes.XlNumber:
    if not float(number).is_integer():
        raise NumExcelError
    
    if number < 0:
        raise NumExcelError
    
    number = int(number)
    
    as_str = str(number)
    permitted_digits = set("01234567")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 8)
    if not (0 <= value < 2**30):
        raise NumExcelError
    
    bit_width = 30
    mask = 1 << bit_width - 1
    
    return (value & ~mask) - (value & mask)

@xl.register()
@xl.validate_args
def OCT2HEX(
    number: func_xltypes.XlNumber,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if places is UNUSED:
        places = None
    else:
        places = int(places)
        # if not (1 <= places <= 10):
        if not (1 <= places <= 10):
            raise NumExcelError
    
    if not float(number).is_integer():
        raise NumExcelError
    
    if number < 0:
        raise NumExcelError
    
    number = int(number)
    as_str = str(number)
    permitted_digits = set("01234567")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 8)
    
    if value >= 2**30: # doublecheck later
        raise NumExcelError
    
    oct_width = 30
    hex_width = 40
    
    mask = 1 << oct_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    negative = new_value < 0
    if negative:
        new_value += (1 << hex_width)
    
    string = hex(new_value)[2:].upper()
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(desired_length)


@xl.register()
@xl.validate_args
def HEX2BIN(
    number: func_xltypes.XlText,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:
    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    as_str = str(number) if number else "0"
    
    try:
        value = int(as_str, 16)
    except ValueError:
        as_float = float(number)
        if not as_float.is_integer():
            raise NumExcelError
        value = int(str(int(as_float)), 16) # oh dear
    
    hex_width = 40
    bin_width = 10
    
    mask = 1 << hex_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    if not (-512 <= new_value < 512):
        raise NumExcelError
    
    negative = new_value < 0
    if negative:
        new_value += (1 << bin_width)
    
    string = bin(new_value)[2:]
    
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)


@xl.register()
@xl.validate_args
def HEX2OCT(
    number: func_xltypes.XlAnything,
    places: func_xltypes.XlNumber = UNUSED
) -> func_xltypes.XlText:    
    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError
    
    if isinstance(number, func_xltypes.Blank):
        as_str = "0"
    
    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))
        
    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)
    
    value = int(as_str, 16)
       
    hex_width = 40
    oct_width = 30
    
    mask = 1 << hex_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    if not (-2**29 <= new_value < 2**29):
        raise NumExcelError
    
    negative = new_value < 0
    if negative:
        new_value += (1 << oct_width)
    
    string = oct(new_value)[2:]
    if places is None:
        return string
    
    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError
    
    return string.zfill(places)


@xl.register()
@xl.validate_args
def HEX2DEC(number: func_xltypes.XlAnything) -> func_xltypes.XlNumber:
    if not number: # covers Blank and empty string
        as_str = "0"
        
    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)
        
    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))
        
    permitted_digits = set("0123456789ABCDEFabcdef")
    if set(as_str) - permitted_digits:
        raise NumExcelError
    
    value = int(as_str, 16)
    if not (0 <= value < 2**40):
        raise NumExcelError
    
    bit_width = 40
    mask = 1 << bit_width - 1
    
    return (value & ~mask) - (value & mask)


# Base = Literal[bin, oct, hex]
# bit_widths = {bin: 10, oct: 30, hex: 40}


# def dec_to_base(value: int, base: Base) -> str:
#     bit_width = bit_widths[base]
#     offset = bit_width - 10
    
#     if value < 0:
#         value += (1 << bit_width)
    
#     if value < 0 and value.bit_length() == 10:
#         value += ((1 << offset) - 1) << 10
    
#     return base(value).strip("-")[2:].upper()


# def handle_places(func):
#     """Handle places awkwardness"""
    
#     def new_func(number: func_xltypes.XlText, places: Optional[func_xltypes.XlNumber] = None) -> func_xltypes.XlText:
#         # print(repr(places))
        
#         if places is not None:
#             places = int(places)
#             if not (1 <= places <= 10):
#                 raise NumExcelError(f"Places must be between 1 and 10; got {places}")
        
#         result = func(number)
        
#         if places is None or str(number).startswith("-"):
#             # negative numbers ignore valid place parameters for some reason
#             return result
        
#         if places < len(result):
#             raise NumExcelError(f"{places} places is not enough to represent {result}")
        
#         return result.zfill(places)
    
#     new_func.__name__ = func.__name__ # can't use functools.wraps bc it breaks the signature
#     return new_func


# def override_int_handling(func):
#     @wraps(func)
#     def new_func(number: func_xltypes.XlText) -> func_xltypes.XlText:
#         try:
#             int(number)
#         except xlerrors.ValueExcelError as e:
#             raise NumExcelError from e
        
#         return func(number)
    
#     return new_func


# def base_to_dec(value: str, base: Base) -> int:
#     value = int(value, {bin: 2, oct: 8, hex: 16}[base])
#     bit_width = bit_widths[base]
#     mask = 1 << bit_width - 1
    
#     return (value & ~mask) - (value & mask)


# def base_to_base(value: str, base_in: Base, base_out: Base) -> str:
#     return dec_to_base(base_to_dec(value, base_in), base_out)


# def parse_number(number: func_xltypes.XlText, base: Base) -> str:
#     as_str = str(number)
    
#     if len(as_str) > 10:
#         raise NumExcelError(f"Input must not have more than 10 hex digits; got {len(as_str)}")
    
#     digits = {bin: "01", oct: "01234567", hex: "0123456789ABCDEF"}[base]
#     if set(as_str.upper()) - set(digits):
#         raise NumExcelError(f"Input must be positive and only contain digits {digits}; got {as_str}")
    
#     return as_str


# @xl.register()
# @xl.validate_args
# @override_int_handling
# def BIN2DEC(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, bin) 
#     return base_to_dec(parsed, bin)


# @xl.register()
# @xl.validate_args
# @handle_places
# @override_int_handling
# def BIN2OCT(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, bin)
#     return base_to_base(parsed, bin, oct)


# @xl.register()
# @xl.validate_args
# @handle_places
# @override_int_handling
# def BIN2HEX(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, bin)
#     return base_to_base(parsed, bin, hex)


# @xl.register()
# @xl.validate_args
# @handle_places
# def DEC2BIN(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     number = int(number) + 1
#     if not (-2**9 <= number < 2**9):
#         raise NumExcelError
    
#     return dec_to_base(number, bin)


# @xl.register()
# @xl.validate_args
# @handle_places
# def DEC2OCT(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     number = int(number)
#     if not (-2**29 <= number < 2**29):
#         raise NumExcelError
    
#     return dec_to_base(number, oct)


# @xl.register()
# @xl.validate_args
# @handle_places
# def DEC2HEX(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     number = int(number)
#     # Note: in LibreOffice Calc the bounds may be different; see
#     # https://bugs.documentfoundation.org/show_bug.cgi?id=139173
#     if not (-2**39 <= number < 2**39):
#         raise NumExcelError
    
#     return dec_to_base(number, hex)
            

# @xl.register()
# @xl.validate_args
# @handle_places
# @override_int_handling
# def OCT2BIN(number: func_xltypes.XlText) -> func_xltypes.XlText:    
#     number = number or "0"
#     if 1000 <= int(number) < 7777777000:
#         raise NumExcelError
    
#     parsed = parse_number(number, oct)
#     return base_to_base(parsed, oct, bin)


# @xl.register()
# @xl.validate_args
# @override_int_handling
# def OCT2DEC(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, oct)
#     return base_to_dec(parsed, oct)


# @xl.register()
# @xl.validate_args
# @handle_places
# @override_int_handling
# def OCT2HEX(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, oct)
#     return base_to_base(parsed, oct, hex)


# @xl.register()
# @xl.validate_args
# @handle_places
# def HEX2BIN(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, hex)
#     if 0x200 <= int(str(parsed), 16) < 0xfffffffe00:
#         raise NumExcelError
    
#     return base_to_base(parsed, hex, bin)


# @xl.register()
# @xl.validate_args
# @handle_places
# def HEX2OCT(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, hex)
#     if 0x20000000 <= int(str(parsed), 16) < 0xffe0000000:
#         raise NumExcelError
    
#     return base_to_base(parsed, hex, oct)


# @xl.register()
# @xl.validate_args
# def HEX2DEC(number: func_xltypes.XlText) -> func_xltypes.XlText:
#     number = number or "0"
#     parsed = parse_number(number, hex)
#     return base_to_dec(parsed, hex)


# # currently breaks with $A5 references