import datetime


def get_YYYYmmdd_cn(date_str_by_dot: str):
    """

    :param date_str_by_dot:2023.06.07
    :return:
    """
    try:
        return datetime.datetime.strptime(date_str_by_dot, '%Y.%m.%d').strftime(
            '%Y年%m月%d日'.encode('unicode_escape').decode('utf8')).encode('utf-8').decode('unicode_escape')
    except BaseException as e:
        print(f'时间有误:{date_str_by_dot}')
        print(e)
        return None


def parse_dot_date_time(date_str_by_dot: str):
    if date_str_by_dot is None:
        return datetime.datetime.strptime('1.1.1', '%Y.%m.%d')
    try:
        return datetime.datetime.strptime(date_str_by_dot, '%Y.%m.%d')
    except BaseException as e:
        print(f'时间有误:{date_str_by_dot}')
        print(e)
        return datetime.datetime.strptime('1.1.1', '%Y.%m.%d')
