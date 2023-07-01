import datetime
from typing import List


def zip_arr(arr: list) -> List[tuple]:
    """
    :return [(start1, end1),(start2, end2),...,(startN, endN)] end is excluded
    """
    if arr is None or len(arr) == 0:
        return []

    if len(arr) == 1:
        return [(0, 1)]
    pre = 0
    ans = []

    for i in range(1, len(arr)):
        if arr[i] is not arr[i - 1]:
            ans.append((pre, i))
            pre = i

    ans.append((pre, len(arr)))

    return ans


def get_one_value(data_list: List[dict], key):
    if data_list is None:
        return None
    for data in data_list:
        if key in data:
            return data[key]
    return None


def sort_date_array(array: List[str], reverse: bool) -> List[str]:
    if array is None:
        return None
    return sorted(array, key=lambda date_str: datetime.datetime.strptime(date_str, "%Y.%m.%d"), reverse=reverse)


if __name__ == '__main__':
    arr = [None, '2021.01.01']
    print(sort_date_array(arr, True))