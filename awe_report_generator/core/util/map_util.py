
def get(map:dict, key:str, default_value=None):
    if map is None:
        return None
    return map.get(key, default_value)
