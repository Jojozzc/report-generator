
def wrap_str(obj):
    if obj is None or str(obj) == 'nan' or str(obj) == 'None':
        return None
    return str(obj)


def wrap_int(obj):
    if str(obj).isdigit():
        return int(obj)
    if is_float(obj):
        return int(obj)
    return None


def wrap_int_str(obj):
    if obj is None or str(obj) == 'nan' or str(obj) == 'None':
        return None
    if is_float(obj):
        return str(int(obj))
    return str(obj)


def is_float(obj):
    try:
        if obj is None or str(obj) == 'nan' or str(obj) == 'None':
            return False
        float(obj)
    except:
        return False
    return True