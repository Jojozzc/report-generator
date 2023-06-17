
def excelTitleToIndex(columnTitle: str) -> int:
    '''
    :param columnTitle:excel column title:A, B, C,... AB,AC,...
    :return: col index:start from 0
    '''
    index, multiple = 0, 1
    for i in range(len(columnTitle) - 1, -1, -1):
        k = ord(columnTitle[i]) - ord("A") + 1
        index += k * multiple
        multiple *= 26
    return index - 1
