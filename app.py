import os, sys

from PyQt5.QtWidgets import QApplication

from awe_report_generator.biz.rt.base import RTHeaderResource
from awe_report_generator.core.base import ExcelFiledProperty, DocGlobalParamConfig
from awe_report_generator.core.simple_table_doc_generator import (
    SimpleTableDocGenerator, MappingColumnCellsResource, SimpleCalculationColumnCellsResource,
    DataListMappingCellResource, GlobalParamMappingCellResource, DataListCellParagraphResource, HeaderResource
)
from awe_report_generator.core.style import DocCellStyle
from awe_report_generator.core.util import cast_util, date_util as awe_date_util

from awe_report_generator.ui import BaseUIConfig
from awe_report_generator.ui.base import HomePageQWidget
from awe_report_generator.ui.processor_ui import ProcessorQWidget, ProcessorUIParam

from awe_report_generator.biz.rt.base import RTSummaryCellResource


def build_rt_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号'),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_int, 'γ射线'),
    }
    column_cell_resource_list = [
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=0,
                                   mapping_data_key='sampleNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=1,
                                   mapping_data_key='kindNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=2,
                                   mapping_data_key='material'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=3,
                                   mapping_data_key='specification'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=4,
                                   mapping_data_key='baseSpecificationAndCnt'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=5,
                                   mapping_data_key='okCount'),
        SimpleCalculationColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=6,
                                             mapping_data_key_1='checkCount', mapping_data_key_2='okCount',
                                             operation='-'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
    ]



    cell_resource_list = [
        DataListMappingCellResource(table_index=0, row=0, column_view_index=3, mapping_key=divide_key),
        DataListMappingCellResource(table_index=0, row=1, column_view_index=3, mapping_key='completeDate'),
        GlobalParamMappingCellResource(table_index=0, row=2, column_view_index=3, mapping_key='testingStandards'),
        DataListMappingCellResource(table_index=0, row=2, column_view_index=5, mapping_key='level'),
        RTSummaryCellResource(table_index=0, row=21, column_view_index=0, style=DocCellStyle(center=False)),
        DataListCellParagraphResource(table_index=0, row=22, column_view_index=3, style=DocCellStyle(right=True), paragraph_index=3, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='completeDate'),
    ]
    template_path = os.path.join(os.path.dirname(__file__), 'template/rt/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list)
    processor_widget = ProcessorQWidget(ProcessorUIParam(title='射线结果生成器', processor=report_generator))
    base_ui_config = BaseUIConfig('射线结果生成器', processor_widget)

    return base_ui_config


def build_ray_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号'),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_int, 'γ射线'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "weldMethod": ExcelFiledProperty('R', cast_util.wrap_str, '焊接方法'),
        "lineNo": ExcelFiledProperty('T', cast_util.wrap_str, '单线号'),

    }

    
    column_cell_resource_list = [
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=1,
                                   mapping_data_key='sampleNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=2,
                                   mapping_data_key='lineNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=3,
                                   mapping_data_key='kindNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=4,
                                   mapping_data_key='empId'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=5,
                                   mapping_data_key='specification'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=7, column_view_index=6,
                                   mapping_data_key='material'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('customorCompany', None, '委托单位', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
        DocGlobalParamConfig('detectionOpportunity', None, '检测时机', False),
        DocGlobalParamConfig('acceptStandard', None, '验收规范', False),
        DocGlobalParamConfig('detectionTechLevel', None, '检测技术等级', False),
        DocGlobalParamConfig('appearanceDetection', '合格', '外观检查', False),
    ]

    def cast_unit_value_to_str(unit_name):
        if unit_name == None:
            return None
        return f'单位工程名称:{unit_name}'


    cell_resource_list = [
        # 单位工程名称
        DataListCellParagraphResource(table_index=0, row=0, column_view_index=2, style=DocCellStyle(center=False ,left=True, right=False), paragraph_index=2, cast_value_to_str=cast_unit_value_to_str, mapping_key='unitName'),
         # 委托单位
        GlobalParamMappingCellResource(table_index=0, row=1, column_view_index=1, mapping_key='customorCompany'),
        # 委托单编号
        DataListMappingCellResource(table_index=0, row=1, column_view_index=3, mapping_key=divide_key),
        # 验收标准
        GlobalParamMappingCellResource(table_index=0, row=2, column_view_index=5, mapping_key='acceptStandard'),
        # 检测标准
        GlobalParamMappingCellResource(table_index=0, row=3, column_view_index=1, mapping_key='testingStandards'),
        # 检测时机
        GlobalParamMappingCellResource(table_index=0, row=3, column_view_index=5, mapping_key='detectionOpportunity'),
        # 检测技术等级
        GlobalParamMappingCellResource(table_index=0, row=4, column_view_index=1, mapping_key='detectionTechLevel'),
        # 检测比列
        DataListMappingCellResource(table_index=0, row=4, column_view_index=3, mapping_key='checkRatioKind'),
        # 合格级别
        DataListMappingCellResource(table_index=0, row=4, column_view_index=5, mapping_key='level'),
        # 焊接方法
        DataListMappingCellResource(table_index=0, row=5, column_view_index=1, mapping_key='weldMethod'),
        # 外观检查
        GlobalParamMappingCellResource(table_index=0, row=5, column_view_index=3, mapping_key='appearanceDetection'),
        # 施工单位时间
        DataListCellParagraphResource(table_index=0, row=23, column_view_index=0, style=DocCellStyle(right=True), paragraph_index=3, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate'),

    ]
    template_path = os.path.join(os.path.dirname(__file__), 'template/ray/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=HeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list)

    processor_widget = ProcessorQWidget(ProcessorUIParam(title='管道焊口检测委托单', processor=report_generator))

    base_ui_config = BaseUIConfig('管道焊口检测委托单', processor_widget)

    return base_ui_config


def build_surface_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号'),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_int, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/m2/点)'),

    }
    column_cell_resource_list = [
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=0,
                                   mapping_data_key='sampleNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=1,
                                   mapping_data_key='kindNo'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=2,
                                   mapping_data_key='material'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=3,
                                   mapping_data_key='specification'),
        MappingColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=5,
                                   mapping_data_key='okCount'),
        SimpleCalculationColumnCellsResource(table_index=0, table_data_start_row=5, column_view_index=6,
                                             mapping_data_key_1='checkCount', mapping_data_key_2='okCount',
                                             operation='-'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('customorCompany', None, '委托单位', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
    ]



    cell_resource_list = [
        DataListMappingCellResource(table_index=0, row=0, column_view_index=3, mapping_key=divide_key),
        DataListMappingCellResource(table_index=0, row=1, column_view_index=3, mapping_key='completeDate'),
        GlobalParamMappingCellResource(table_index=0, row=2, column_view_index=3, mapping_key='testingStandards'),
        DataListMappingCellResource(table_index=0, row=2, column_view_index=5, mapping_key='level'),
        RTSummaryCellResource(table_index=0, row=21, column_view_index=0, style=DocCellStyle(center=False)),
        DataListCellParagraphResource(table_index=0, row=22, column_view_index=3, style=DocCellStyle(right=True), paragraph_index=3, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='completeDate'),
    ]
    template_path = os.path.join(os.path.dirname(__file__), 'template/surface/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=HeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list)
    processor_widget = ProcessorQWidget(ProcessorUIParam(title='表面结果生成器', processor=report_generator))
    base_ui_config = BaseUIConfig('表面结果生成器', processor_widget)

    return base_ui_config


if __name__ == '__main__':
    # Must run before QWidgets init.
    app = QApplication(sys.argv)

    ui_config_list = [
        build_ray_config(),
        build_rt_config(),
        build_surface_config(),
    ]
    ui = HomePageQWidget(ui_config_list)
    ui.show()

    sys.exit(app.exec_())
