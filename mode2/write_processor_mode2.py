import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
import pandas
import os
from mode2.RawDataMode2 import RawDataMode2
import datetime

TABLE_MAPPING = {
    "complete_date": 1,
    "order_id": 2,
    "sample_no": 3,
    "line_no": 4,
    "kind_no": 5,
    "emp_id": 6,
    "material": 8,
    "specification": 7,
    "ret_cnt": 11,
    "qualified_sample_cnt": 12,
}


def write_cell(cell, text: str):
    if text is None:
        text = ''
    cell.text = text
    cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

def write_cell_para(cell, para_idx: int,text: str):
    if cell.paragraphs is None:
        print(f'模板格式有误，请检查，待输入文字:{text}')
        return
    cell.paragraphs[para_idx].text = text


'''
:param callback (total_cnt, fin_cnt, fail_cnt)
'''
def process(template_path: str, input_file_path: str, title1: str, title2: str, customer: str,
            method: str, target_dir: str, callback):
    base_start = 3
    raw_data_map = read_raw_data(input_file_path)

    total_cnt = len(raw_data_map.keys())
    fin_cnt = 0
    fail_cnt = 0
    problem_cnt = 0
    callback(total_cnt, fin_cnt, fail_cnt, problem_cnt)

    for key in raw_data_map.keys():
        has_problem = False
        try:
            callback(total_cnt, fin_cnt, fail_cnt, problem_cnt)
            values = raw_data_map[key]
            order_id = key
            complete_date = values[0].complete_date

            doc = docx.Document(template_path)

            target_path = os.path.join(target_dir, f'{order_id}.docx')

            table = doc.tables[0]
            write_cell_para(table.cell(0, 12), 0, f"工程名称:{title1}\n单位工程名称:{title2}")

            write_cell(table.cell(1, 12), order_id)  # 通知单号

            write_cell(table.cell(1, 2), customer)  # 委托单位

            write_cell(table.cell(1, 6), method)  # 检测方法


            date_cn = get_YYYYmmdd_cn(complete_date)

            last_idx = len(table.rows) - 1
            if date_cn is not None:
                write_cell_para(table.cell(last_idx, 1), 3, f'日期：  {date_cn}')
                write_cell_para(table.cell(last_idx, 8), 3, f'日期：  {date_cn}')
            else:
                has_problem = True
                print(f'日期有误，通知单号:{key}')


            # 数据从第4行开始写入
            for i in range(len(values)):
                raw_data: RawDataMode2 = values[i]

                # 委托编号
                write_cell(table.cell(i + base_start, 1), raw_data.sample_no)

                # 检测批号
                write_cell(table.cell(i + base_start, 3), '/')

                # 单线号
                write_cell(table.cell(i + base_start, 4), raw_data.line_no)

                # 焊口编号
                write_cell(table.cell(i + base_start, 9), raw_data.kind_no)

                # 焊工号
                write_cell(table.cell(i + base_start, 12), raw_data.emp_id)

                # 返修张/处数
                ret_cnt_str = wrap_str(raw_data.ret_cnt)
                if ret_cnt_str is None:
                    ret_cnt_str = '/'
                write_cell(table.cell(i + base_start, 13), wrap_str(raw_data.ret_cnt))
                write_cell(table.cell(i + base_start, 13), ret_cnt_str)

            doc.save(target_path)
            if has_problem:
                problem_cnt = problem_cnt + 1
        except Exception as e:
            print(f'数据有误,请检查：通知单号={key}')
            print(e)
            fail_cnt = fail_cnt + 1
        finally:
            fin_cnt = fin_cnt + 1

    callback(total_cnt, fin_cnt, fail_cnt, problem_cnt)



def read_raw_data(input_file_path: str):
    raw_datas = pandas.read_excel(input_file_path, 'Sheet1')
    data_map = {}

    for i in range(raw_datas.shape[0]):
        row = raw_datas.iloc[i]
        order_id = wrap_int_str(row[TABLE_MAPPING['order_id']])

        if order_id is None:
            print(f'非法的通知单号:${order_id},请检查')
            continue

        if order_id not in data_map:
            data_map[order_id] = []

        raw_data = RawDataMode2(complete_date=wrap_str(row[TABLE_MAPPING['complete_date']]),
                           sample_no=wrap_str(row[TABLE_MAPPING['sample_no']]),
                           line_no=wrap_str(row[TABLE_MAPPING['line_no']]),
                           kind_no=wrap_str(row[TABLE_MAPPING['kind_no']]),
                           emp_id=wrap_str(row[TABLE_MAPPING['emp_id']]),
                           ret_cnt=wrap_int(row[TABLE_MAPPING['ret_cnt']]),
                           material=wrap_str(row[TABLE_MAPPING['material']]),
                           specification=wrap_str(row[TABLE_MAPPING['specification']]),
                           qualified_sample_cnt=wrap_int(row[TABLE_MAPPING['qualified_sample_cnt']]))
        data_map[order_id].append(raw_data)

    return data_map


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


def get_YYYYmmdd_cn(date_str_by_dot: str):
    try:
        return datetime.datetime.strptime(date_str_by_dot, '%Y.%m.%d').strftime('%Y年%m月%d日'.encode('unicode_escape').decode('utf8')).encode('utf-8').decode('unicode_escape')
    except BaseException as e:
        print(f'时间有误:{date_str_by_dot}')
        print(e)
        return None