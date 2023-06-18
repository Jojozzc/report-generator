import os, sys

from PyQt5.QtWidgets import QApplication

from awe_report_generator.biz.rt.base import RTHeaderResource
from awe_report_generator.core.base import ExcelFiledProperty, DocGlobalParamConfig
from awe_report_generator.core.simple_table_doc_generator import (
    SimpleTableDocGenerator, MappingColumnCellsResource, SimpleCalculationColumnCellsResource,
    DataListMappingCellResource
)
from awe_report_generator.core.util import cast_util
from awe_report_generator.ui.base import HomePageQWidget, ProcessorUIConfig


def build_rt_config():
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
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=0, mapping_data_key='sampleNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=2, mapping_data_key='kindNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=4, mapping_data_key='material'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=5, mapping_data_key='specification'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=9,
                                   mapping_data_key='baseSpecificationAndCnt'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=12, mapping_data_key='okCount'),
        SimpleCalculationColumnCellsResource(table_index=0, table_data_start_row=5, column=14,
                                             mapping_data_key_1='checkCount', mapping_data_key_2='okCount',
                                             operation='-'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', True),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
    ]

    cell_resource_list = [
        DataListMappingCellResource(table_index=0, row=0, column=14, mapping_key=divide_key),
        DataListMappingCellResource(table_index=0, row=1, column=14, mapping_key='completeDate'),
        DataListMappingCellResource(table_index=0, row=2, column=14, mapping_key='level'),
    ]
    template_path = os.path.join(os.path.dirname(__file__), 'template/rt/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list)

    return ProcessorUIConfig(name='RT', processor=report_generator)


def build_ray_config():
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
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=0, mapping_data_key='sampleNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=2, mapping_data_key='kindNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=4, mapping_data_key='material'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=5, mapping_data_key='specification'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=9,
                                   mapping_data_key='baseSpecificationAndCnt'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column=12, mapping_data_key='okCount'),
        SimpleCalculationColumnCellsResource(table_index=0, table_data_start_row=5, column=14,
                                             mapping_data_key_1='checkCount', mapping_data_key_2='okCount',
                                             operation='-'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
    ]

    cell_resource_list = [
        DataListMappingCellResource(table_index=0, row=0, column=14, mapping_key=divide_key),
        DataListMappingCellResource(table_index=0, row=1, column=14, mapping_key='completeDate'),
        DataListMappingCellResource(table_index=0, row=2, column=14, mapping_key='level'),
    ]
    template_path = os.path.join(os.path.dirname(__file__), 'template/rt/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list)

    return ProcessorUIConfig(name='RAY', processor=report_generator)

if __name__ == '__main__':
    processor_ui_list = [
        build_rt_config(),
        build_ray_config(),
    ]
    app = QApplication(sys.argv)
    ui = HomePageQWidget(processor_ui_list)
    ui.show()

    sys.exit(app.exec_())