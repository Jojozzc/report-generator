def wrap_str(obj):
    if obj is None or str(obj) == 'nan' or str(obj) == 'None':
        return None
    return str(obj)


def wrap_int(obj):
    if str(obj).isdigit():
        return int(obj)
    if _is_float(obj):
        return int(obj)
    return None


def wrap_int_str(obj):
    if obj is None or str(obj) == 'nan' or str(obj) == 'None':
        return None
    if _is_float(obj):
        return str(int(obj))
    return str(obj)


def _is_float(obj):
    try:
        if obj is None or str(obj) == 'nan' or str(obj) == 'None':
            return False
        float(obj)
    except:
        return False
    return True


def wrap_float(obj):
    if _is_float(obj):
        return float(obj)
    return None


def wrap_percent(obj):
    if obj is None:
        return None
    try:
        number = wrap_float(obj)
        return "{:.0%}".format(number)
    except BaseException as e:
        return str(obj)


if __name__ == '__main__':
    print(wrap_percent('0.1'))