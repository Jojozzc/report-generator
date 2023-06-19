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
        
