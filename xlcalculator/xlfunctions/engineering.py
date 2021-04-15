from . import xl, func_xltypes
from .xlerrors import NumExcelError, ValueExcelError


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


@xl.register()
@xl.validate_args
def DEC2BIN(
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    number = int(number)
    if not (-513 < number < 512):
        raise NumExcelError

    negative = number < 0
    bit_width = 10
    if negative:
        number += 1 << bit_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    number = int(number)

    if not (-(2 ** 29) <= number < 2 ** 29):
        raise NumExcelError

    bit_width = 30
    negative = number < 0

    if negative:
        number += 1 << bit_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    number = int(number)

    if not (-(2 ** 39) <= number < 2 ** 39):
        raise NumExcelError

    bit_width = 40
    negative = number < 0

    if negative:
        number += 1 << bit_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
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
    
    if not (-2 ** 9 <= new_value < 2 ** 9):
        raise NumExcelError

    negative = new_value < 0
    if negative:
        new_value += 1 << oct_width

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
def BIN2DEC(number: func_xltypes.XlAnything) -> func_xltypes.XlNumber:
    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)
        
    if set(as_str) - set("01"):
        raise NumExcelError

    bit_width = 10
    mask = 1 << bit_width - 1

    value = int(as_str, 2)
    new_value = (value & ~mask) - (value & mask)
    
    if not (-2**9 <= new_value < 2 ** 9):
        raise NumExcelError
    
    return new_value 


@xl.register()
@xl.validate_args
def BIN2HEX(
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)

    permitted_digits = set("01")
    if set(as_str) - permitted_digits:
        raise NumExcelError

    value = int(as_str, 2)

    bin_width = 10
    hex_width = 40

    mask = 1 << bin_width - 1
    new_value = (value & ~mask) - (value & mask)

    if not (-(2 ** 9) <= new_value < 2 ** 9):
        raise NumExcelError

    negative = new_value < 0
    if negative:
        new_value += 1 << hex_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
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
        new_value += 1 << bin_width

    string = bin(new_value)[2:]
    if places is None:
        return string

    desired_length = len(string) if negative else places
    if desired_length < len(string):
        raise NumExcelError

    return string.zfill(places)


@xl.register()
@xl.validate_args
def OCT2DEC(number: func_xltypes.XlAnything) -> func_xltypes.XlNumber:
    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)

    permitted_digits = set("01234567")
    if set(as_str) - permitted_digits:
        raise NumExcelError

    value = int(as_str, 8)
    if not (0 <= value < 2 ** 30):
        raise NumExcelError

    bit_width = 30
    mask = 1 << bit_width - 1

    return (value & ~mask) - (value & mask)


@xl.register()
@xl.validate_args
def OCT2HEX(
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        # if not (1 <= places <= 10):
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)

    permitted_digits = set("01234567")
    if set(as_str) - permitted_digits:
        raise NumExcelError

    value = int(as_str, 8)

    oct_width = 30
    hex_width = 40

    mask = 1 << oct_width - 1
    new_value = (value & ~mask) - (value & mask)
    
    if not (-2 ** 30 <= new_value < 2 ** 30):
        raise NumExcelError

    negative = new_value < 0
    if negative:
        new_value += 1 << hex_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)
        
    permitted_digits = set("0123456789ABCDEFabcdef")
    if set(as_str) - permitted_digits:
        raise NumExcelError

    value = int(as_str, 16)

    hex_width = 40
    bin_width = 10

    mask = 1 << hex_width - 1
    new_value = (value & ~mask) - (value & mask)

    if not (-2 ** 9 <= new_value < 2 ** 9):
        raise NumExcelError

    negative = new_value < 0
    if negative:
        new_value += 1 << bin_width

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
    number: func_xltypes.XlAnything, places: func_xltypes.XlAnything = UNUSED
) -> func_xltypes.XlText:
    if isinstance(places, func_xltypes.Boolean):
        raise ValueExcelError

    if places is UNUSED:
        places = None
    else:
        places = int(places)
        if not (1 <= places <= 10):
            raise NumExcelError

    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if isinstance(number, func_xltypes.Blank):
        as_str = "0"

    elif isinstance(number, func_xltypes.Number):
        if number.is_decimal and not number.value.is_integer():
            raise NumExcelError
        as_str = str(int(number))

    elif isinstance(number, func_xltypes.Text):
        as_str = str(number)
    
    permitted_digits = set("0123456789ABCDEFabcdef")
    if set(as_str) - permitted_digits:
        raise NumExcelError

    value = int(as_str, 16)

    hex_width = 40
    oct_width = 30

    mask = 1 << hex_width - 1
    new_value = (value & ~mask) - (value & mask)

    if not (-(2 ** 29) <= new_value < 2 ** 29):
        raise NumExcelError

    negative = new_value < 0
    if negative:
        new_value += 1 << oct_width

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
    if isinstance(number, func_xltypes.Boolean):
        raise ValueExcelError

    if not number:  # covers Blank and empty string
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
    if not (0 <= value < 2 ** 40):
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