import json
import sys

import click
from tqdm import tqdm

from awe_report_generator.biz.radio import (
    RadioCheckGenerator,
)
from awe_report_generator.biz.rt.base import RTHeaderResource
from awe_report_generator.core.base import ExcelFiledProperty, DocGlobalParamConfig
from awe_report_generator.core.simple_table_doc_generator import (
    SimpleTableDocGenerator, MappingColumnCellsParagraphAddRunResource,
    SimpleCalculationColumnCellsParagraphAddRunResource, DataListCellParagraphAddRunResource
)
from awe_report_generator.core.util import cast_util


@click.command()
@click.option('-t', '--template-path')
@click.option('-f', '--file-path')
@click.option('-s', '--sheet')
@click.option('-d', '--target-dir')
@click.option('-m', '--mode')
@click.option('-p', '--global-param')
def run(template_path: str, file_path: str, sheet: str, target_dir: str, mode: str, global_param: str):
    report_generator = None

    pbar = None

    def on_finish_one(number, total_cnt, success, exception):
        nonlocal pbar

        if pbar is None:
            pbar = tqdm(total=total_cnt)

        pbar.update(1)

    global_param_dict = None
    if mode == 'ray':
        report_generator = RadioCheckGenerator(template_path=template_path)
    elif mode == 'rt':
        divide_key = 'orderId'
        filed_mapping = {
            "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
            "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
            divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
            "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
            "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
            "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号'),
            "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格（mm）'),
            "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
            "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
            "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
            "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
            "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
            "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
            "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
            "ray": ExcelFiledProperty('P', cast_util.wrap_int, 'γ射线'),
        }
        column_cell_resource_list = [
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=0, mapping_data_key='sampleNo'),
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=2, mapping_data_key='kindNo'),
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=4, mapping_data_key='material'),
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5, mapping_data_key='specification'),
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=9, mapping_data_key='baseSpecificationAndCnt'),
            MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=12, mapping_data_key='okCount'),
            SimpleCalculationColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=14, mapping_data_key_1='checkCount', mapping_data_key_2='okCount', operation='-'),
        ]

        doc_global_data_param_config_list = [
            DocGlobalParamConfig('projectName', None, '工程名称', False),
            DocGlobalParamConfig('testingStandards', None, '检测标准', False),
        ]

        cell_resource_list = [
            DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=14, mapping_key=divide_key),
            DataListCellParagraphAddRunResource(table_index=0, row=1, column_view_index=14, mapping_key='completeDate'),
            DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=14, mapping_key='level'),
        ]

        global_param_dict = json.loads(global_param)
        report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping, divide_key=divide_key, header_resource=RTHeaderResource(), column_cell_resource_list=column_cell_resource_list, cell_resource_list=cell_resource_list, doc_global_data_param_config_list=doc_global_data_param_config_list)
    if report_generator is None:
        sys.exit(f'Unknown mode:{mode}')

    report_generator.execute(file_path=file_path, target_dir=target_dir, sheet=sheet, global_param=global_param_dict, on_finish_one=on_finish_one)


if __name__ == '__main__':
    run()