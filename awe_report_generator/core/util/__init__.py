
def excel_title_to_index(column_title: str) -> int:
    '''
    :param column_title:excel column title:A, B, C,... AB,AC,...
    :return: col index:start from 0
    '''
    index, multiple = 0, 1
    for i in range(len(column_title) - 1, -1, -1):
        k = ord(column_title[i]) - ord("A") + 1
        index += k * multiple
        multiple *= 26
    return index - 1



