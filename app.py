import os
import sys
from typing import List, Dict

from PyQt5.QtWidgets import QApplication
from docx.enum.table import WD_TABLE_ALIGNMENT

from awe_report_generator.biz.rt.base import RTHeaderResource, RTSummaryCellResource2, RecordHeaderResource, \
    RTRecordSummaryCellParagraphAddRunResource, RayAdditionalHeaderResource
from awe_report_generator.biz.rt.base import RTSummaryCellResource
from awe_report_generator.core.base import ExcelFiledProperty, DocGlobalParamConfig, ConstantValueGetter, \
    DataListMappingValueGetter, GeneratorExecuteContext, FunctionValueGetter
from awe_report_generator.core.merged_excel_generator import MergedExcelReportGenerator, ExcelColumnsValueSetter
from awe_report_generator.core.simple_table_doc_generator import (
    SimpleTableDocGenerator, MappingColumnCellsParagraphAddRunResource,
    SimpleCalculationColumnCellsParagraphAddRunResource,
    HeaderResource,
    GlobalParamCellParagraphAddRunResource,
    DataListCellParagraphAddRunResource, CellParagraphAddRunValueSetter, FuncColumnCellsParagraphAddRunResource
)
from awe_report_generator.core.style import DocCellStyle
from awe_report_generator.core.util import array_util
from awe_report_generator.core.util import cast_util, date_util as awe_date_util
from awe_report_generator.core.util import excel_util
from awe_report_generator.core.util import map_util
from awe_report_generator.ui import BaseUIConfig
from awe_report_generator.ui.base import HomePageQWidget
from awe_report_generator.ui.processor_ui import ProcessorQWidget, ProcessorUIParam


def date_merge_fun(date_list):
    if date_list is None or len(date_list) == 0:
        return None
    date_list = array_util.sort_date_array(date_list, True)
    return date_list[0]

# RT结果通知单
def build_rt_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/㎡/点)'),
    }
    column_cell_resource_list = [
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=0,
                                                  mapping_data_key='sampleNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=1,
                                                  mapping_data_key='kindNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=2,
                                                  mapping_data_key='material', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=3,
                                                  mapping_data_key='specification', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=4,
                                                  mapping_data_key='baseSpecificationAndCnt',
                                                  style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5,
                                                  mapping_data_key='detectionCount',
                                                  style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5,
                                                  mapping_data_key='okCount', style=DocCellStyle(font_cn='楷体')),
        SimpleCalculationColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=6,
                                                            style=DocCellStyle(font_cn='楷体'),
                                                            mapping_data_key_1='checkCount',
                                                            mapping_data_key_2='okCount',
                                                            operation='-'),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
        DocGlobalParamConfig('customCompany', None, '委托单位', False),
        DocGlobalParamConfig('detectionMethod', 'RT', '检测方法', False),
    ]

    cell_resource_list = [
        DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=3, mapping_key=divide_key,
                                            style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=1, column_view_index=1, mapping_key='customCompany',
                                               style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=1, column_view_index=3, mapping_key='completeDate',
                                            style=DocCellStyle(font_cn='楷体'),
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='detectionMethod',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3,
                                               mapping_key='testingStandards', style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=5, mapping_key='level',
                                            style=DocCellStyle(font_cn='楷体')),
        RTSummaryCellResource(table_index=0, row=21, column_view_index=0,
                              style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.LEFT)),
        DataListCellParagraphAddRunResource(table_index=0, row=22, column_view_index=3,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            run_index=None, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),
        DataListCellParagraphAddRunResource(table_index=0, row=22, column_view_index=2,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            run_index=None, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),
        DataListCellParagraphAddRunResource(table_index=0, row=22, column_view_index=1,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            run_index=None, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),
        DataListCellParagraphAddRunResource(table_index=0, row=22, column_view_index=0,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            run_index=None, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),

    ]
    template_path = os.path.join(os.getcwd(), 'template/rt/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list,
                                               merge_fun_dict={'completeDate': date_merge_fun})
    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='RT结果通知单台账', processor=report_generator, biz_code='rt'))
    base_ui_config = BaseUIConfig('RT结果通知单台账', processor_widget)

    return base_ui_config


# 射线检测委托台账
def build_ray_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "weldMethod": ExcelFiledProperty('R', cast_util.wrap_str, '焊接方法'),
        "areaNo": ExcelFiledProperty('S', cast_util.wrap_str, '区号'),
        "lineNo": ExcelFiledProperty('T', cast_util.wrap_str, '单线号'),
        "detectionOpportunity": ExcelFiledProperty('V', cast_util.wrap_str, '检测时机'),

    }

    column_cell_resource_list = [
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=1,
                                                  mapping_data_key='sampleNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=2,
                                                  mapping_data_key='lineNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=3,
                                                  mapping_data_key='kindNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=4,
                                                  mapping_data_key='empId', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=5,
                                                  mapping_data_key='specification',
                                                  style=DocCellStyle(font_cn='楷体', font_size=7.5)),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=7, column_view_index=6,
                                                  mapping_data_key='material', style=DocCellStyle(font_cn='楷体')),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('customorCompany', None, '委托单位', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
        DocGlobalParamConfig('acceptStandard', None, '验收规范', False),
        DocGlobalParamConfig('detectionMethod', None, '检测方法', False),
        DocGlobalParamConfig('detectionTechLevel', None, '检测技术等级', False),
        DocGlobalParamConfig('appearanceDetection', '合格', '外观检查', False),
        DocGlobalParamConfig('groove', 'V', '坡口形式', False),
    ]

    cell_resource_list = [
        # 工程名称
        GlobalParamCellParagraphAddRunResource(table_index=0, row=0, column_view_index=2,
                                               style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.LEFT, font_cn='楷体'),
                                               paragraph_index=0, cast_value_to_str=cast_util.wrap_str,
                                               mapping_key='projectName'),
        # 单位工程名称
        DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=2,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.LEFT, font_cn='楷体'),
                                            paragraph_index=1, cast_value_to_str=cast_util.wrap_str,
                                            mapping_key='unitName'),
        # 委托单位
        GlobalParamCellParagraphAddRunResource(table_index=0, row=1, column_view_index=1, mapping_key='customorCompany',
                                               style=DocCellStyle(font_cn='楷体')),
        # 委托单编号
        DataListCellParagraphAddRunResource(table_index=0, row=1, column_view_index=3, mapping_key=divide_key,
                                            style=DocCellStyle(font_cn='楷体')),
        # 区号
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3, mapping_key='areaNo',
                                            style=DocCellStyle(font_cn='楷体')),
        # 验收规范
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=5, mapping_key='acceptStandard',
                                               style=DocCellStyle(font_cn='楷体')),
        # 检测标准
        GlobalParamCellParagraphAddRunResource(table_index=0, row=3, column_view_index=1,
                                               mapping_key='testingStandards', style=DocCellStyle(font_cn='楷体')),
        # 检测方法
        GlobalParamCellParagraphAddRunResource(table_index=0, row=3, column_view_index=3, mapping_key='detectionMethod',
                                               style=DocCellStyle(font_cn='楷体')),
        # 检测时机
        DataListCellParagraphAddRunResource(table_index=0, row=3, column_view_index=5,
                                            mapping_key='detectionOpportunity', style=DocCellStyle(font_cn='楷体')),
        # 检测技术等级
        GlobalParamCellParagraphAddRunResource(table_index=0, row=4, column_view_index=1,
                                               mapping_key='detectionTechLevel', style=DocCellStyle(font_cn='楷体')),
        # 检测比例
        DataListCellParagraphAddRunResource(table_index=0, row=4, column_view_index=3, mapping_key='checkRatioKind',
                                            style=DocCellStyle(font_cn='楷体')),
        # 合格级别
        DataListCellParagraphAddRunResource(table_index=0, row=4, column_view_index=5, mapping_key='level',
                                            style=DocCellStyle(font_cn='楷体')),
        # 焊接方法
        DataListCellParagraphAddRunResource(table_index=0, row=5, column_view_index=1, mapping_key='weldMethod',
                                            style=DocCellStyle(font_cn='楷体')),
        # 外观检查
        GlobalParamCellParagraphAddRunResource(table_index=0, row=5, column_view_index=3,
                                               mapping_key='appearanceDetection', style=DocCellStyle(font_cn='楷体')),
        # 施工单位时间
        DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=0,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),

        # 监理单位时间
        DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=1,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),

        # 项目部/装置时间
        DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=2,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),

        # 检测单位：（签章）时间
        DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=3,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate',
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),


        # 坡口形式
        GlobalParamCellParagraphAddRunResource(table_index=0, row=5, column_view_index=5, mapping_key='groove',
                                               style=DocCellStyle(font_cn='楷体')),
    ]

    template_path = os.path.join(os.getcwd(), 'template/ray/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=HeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list,
                                               merge_fun_dict={'completeDate': date_merge_fun})

    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='射线检测委托台账', processor=report_generator, biz_code='ray'))

    base_ui_config = BaseUIConfig('射线检测委托台账', processor_widget)

    return base_ui_config

# 表面结果通知单
def build_surface_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/m2/点)'),

    }

    def ok_func(data: dict, global_data: dict, merged_data: dict):
        if data['isOk'] == '合格':
            return data['detectionCount']
        elif data['isOk'] == '不合格':
            return '0'

    def not_ok_func(data: dict, global_data: dict, merged_data: dict):
        if data['isOk'] == '合格':
            return '0'
        elif data['isOk'] == '不合格':
            return data['detectionCount']

    column_cell_resource_list = [
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=0,
                                                  mapping_data_key='sampleNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=1,
                                                  mapping_data_key='kindNo', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=2,
                                                  mapping_data_key='material', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=3,
                                                  mapping_data_key='specification', style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=4,
                                                  mapping_data_key='detectionCount',
                                                  style=DocCellStyle(font_cn='楷体')),
        # MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5,
        #                                           mapping_data_key='okCount', style=DocCellStyle(font_cn='楷体')),
        # SimpleCalculationColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=6,
        #                                                     mapping_data_key_1='checkCount',
        #                                                     mapping_data_key_2='okCount',
        #                                                     operation='-', style=DocCellStyle(font_cn='楷体')),

        FuncColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5, style=DocCellStyle(font_cn='楷体'), func=ok_func),
        FuncColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=6, style=DocCellStyle(font_cn='楷体'), func=not_ok_func),
    ]

    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('customerCompany', None, '委托单位', False),
        DocGlobalParamConfig('testingStandards', None, '检测标准', False),
        DocGlobalParamConfig('detectionMethod', 'RT', '检测方法', False),
    ]

    cell_resource_list = [
        DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=3, paragraph_index=0,
                                            mapping_key=divide_key, style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=1, column_view_index=3, paragraph_index=0,
                                            mapping_key='completeDate', style=DocCellStyle(font_cn='楷体'),
                                            data_list_sort_func=awe_date_util.parse_dot_date_time,
                                            data_list_sort_reverse=True),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=1, column_view_index=1, mapping_key='customerCompany',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='detectionMethod',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3,
                                               mapping_key='testingStandards', style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=5, paragraph_index=0,
                                            mapping_key='level', style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=25, column_view_index=3,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=4,
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate'),
        RTSummaryCellResource2(table_index=0, row=24, column_view_index=0,
                               style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.LEFT, font_cn='楷体')),
    ]
    template_path = os.path.join(os.getcwd(), 'template/surface/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list,
                                               merge_fun_dict={'completeDate': date_merge_fun})
    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='表面结果通知单台账', processor=report_generator, biz_code='surface'))
    base_ui_config = BaseUIConfig('表面结果通知单台账', processor_widget)

    return base_ui_config


# 质量评定台账excel
def build_record_excel_config():
    divide_key = ''
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        'orderId': ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质', column_type=str),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "weldMethod": ExcelFiledProperty('R', cast_util.wrap_str, '焊接方法'),
        "areaNo": ExcelFiledProperty('S', cast_util.wrap_str, '区号', column_type=str),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/m2/点)'),
        "detectionOpportunity": ExcelFiledProperty('V', cast_util.wrap_str, '检测时机'),

    }

    doc_global_data_param_config_list = [DocGlobalParamConfig('expectDate', None, '生成时间', True, 'date')]

    excel_header_value_setter_list = [
        ExcelColumnsValueSetter(column='A', value_getter=ConstantValueGetter('委托编号')),
        ExcelColumnsValueSetter(column='B', value_getter=ConstantValueGetter('检件编号')),
        ExcelColumnsValueSetter(column='C', value_getter=ConstantValueGetter('焊口号')),
        ExcelColumnsValueSetter(column='D', value_getter=ConstantValueGetter('焊工号')),
        ExcelColumnsValueSetter(column='E', value_getter=ConstantValueGetter('焊口规格')),
        ExcelColumnsValueSetter(column='F', value_getter=ConstantValueGetter('片号')),
        ExcelColumnsValueSetter(column='G', value_getter=ConstantValueGetter('缺陷性质')),
        ExcelColumnsValueSetter(column='H', value_getter=ConstantValueGetter('缺陷定量')),
        ExcelColumnsValueSetter(column='I', value_getter=ConstantValueGetter('评定级别')),
        ExcelColumnsValueSetter(column='J', value_getter=ConstantValueGetter('透照方式')),
        ExcelColumnsValueSetter(column='K', value_getter=ConstantValueGetter('像质计灵敏度')),
        ExcelColumnsValueSetter(column='L', value_getter=ConstantValueGetter('焦距（mm）')),
        ExcelColumnsValueSetter(column='M', value_getter=ConstantValueGetter('有效片长')),
        ExcelColumnsValueSetter(column='N', value_getter=ConstantValueGetter('源强（管电压）')),
        ExcelColumnsValueSetter(column='O', value_getter=ConstantValueGetter('管电流源活度')),
        ExcelColumnsValueSetter(column='P', value_getter=ConstantValueGetter('曝光量时间')),
        ExcelColumnsValueSetter(column='Q', value_getter=ConstantValueGetter('设备型号射源种类')),
        ExcelColumnsValueSetter(column='R', value_getter=ConstantValueGetter('焦点尺寸')),
        ExcelColumnsValueSetter(column='S', value_getter=ConstantValueGetter('增感方式')),
        ExcelColumnsValueSetter(column='T', value_getter=ConstantValueGetter('胶片牌号')),
        ExcelColumnsValueSetter(column='U', value_getter=ConstantValueGetter('备注')),
        ExcelColumnsValueSetter(column='V', value_getter=ConstantValueGetter('区号')),
        ExcelColumnsValueSetter(column='W', value_getter=ConstantValueGetter('合格级别')),
        ExcelColumnsValueSetter(column='X', value_getter=ConstantValueGetter('管道材质')),
        ExcelColumnsValueSetter(column='Y', value_getter=ConstantValueGetter('焊接方法')),
        ExcelColumnsValueSetter(column='Z', value_getter=ConstantValueGetter('检测比例')),
        ExcelColumnsValueSetter(column='AA', value_getter=ConstantValueGetter('检测比例')),
    ]

    excel_columns_value_setter_list = [
        ExcelColumnsValueSetter(column='A', value_getter=DataListMappingValueGetter('orderId')),
        ExcelColumnsValueSetter(column='B', value_getter=DataListMappingValueGetter('sampleNo')),
        ExcelColumnsValueSetter(column='C', value_getter=DataListMappingValueGetter('kindNo')),
        ExcelColumnsValueSetter(column='D', value_getter=DataListMappingValueGetter('empId')),
        ExcelColumnsValueSetter(column='E', value_getter=DataListMappingValueGetter('specification')),
        ExcelColumnsValueSetter(column='F', value_getter=DataListMappingValueGetter('_pieceNo')),

        ExcelColumnsValueSetter(column='J', value_getter=DataListMappingValueGetter('_guide.B')),
        ExcelColumnsValueSetter(column='K', value_getter=DataListMappingValueGetter('_guide.C')),
        ExcelColumnsValueSetter(column='L', value_getter=DataListMappingValueGetter('_guide.D')),
        ExcelColumnsValueSetter(column='M', value_getter=DataListMappingValueGetter('_guide.E')),
        ExcelColumnsValueSetter(column='N', value_getter=DataListMappingValueGetter('_guide.F')),
        ExcelColumnsValueSetter(column='O', value_getter=DataListMappingValueGetter('_guide.G')),
        ExcelColumnsValueSetter(column='P', value_getter=DataListMappingValueGetter('_guide.H')),
        ExcelColumnsValueSetter(column='Q', value_getter=DataListMappingValueGetter('_guide.J')),
        ExcelColumnsValueSetter(column='R', value_getter=DataListMappingValueGetter('_guide.K')),
        ExcelColumnsValueSetter(column='S', value_getter=DataListMappingValueGetter('_guide.L')),
        ExcelColumnsValueSetter(column='T', value_getter=DataListMappingValueGetter('_guide.M')),

        ExcelColumnsValueSetter(column='V', value_getter=DataListMappingValueGetter('areaNo')),
        ExcelColumnsValueSetter(column='W', value_getter=DataListMappingValueGetter('level')),
        ExcelColumnsValueSetter(column='X', value_getter=DataListMappingValueGetter('material')),
        ExcelColumnsValueSetter(column='Y', value_getter=DataListMappingValueGetter('weldMethod')),
        ExcelColumnsValueSetter(column='Z', value_getter=DataListMappingValueGetter('checkRatioKind')),
        ExcelColumnsValueSetter(column='AA', value_getter=DataListMappingValueGetter('detectionOpportunity')),

    ]

    def data_preparer(data_list: List[dict], context: GeneratorExecuteContext):
        new_data_list = []

        guide_data_map = context.get_data('GUIDE_DATA')

        if guide_data_map is None:
            guide_filed_mapping = {
                "A": ExcelFiledProperty('A', cast_util.wrap_str, '规格', column_type=str),
                "B": ExcelFiledProperty('B', cast_util.wrap_str, '透照方式', column_type=str),
                "C": ExcelFiledProperty('C', cast_util.wrap_str, '像质计灵敏度', column_type=str),
                "D": ExcelFiledProperty('D', cast_util.wrap_str, '焦距（mm）', column_type=str),
                "E": ExcelFiledProperty('E', cast_util.wrap_str, '有效片长', column_type=str),
                "F": ExcelFiledProperty('F', cast_util.wrap_str, '源强（管电压）', column_type=str),
                "G": ExcelFiledProperty('G', cast_util.wrap_str, '管电流源活度', column_type=str),
                "H": ExcelFiledProperty('H', cast_util.wrap_str, '曝光量时间', column_type=str),
                "I": ExcelFiledProperty('I', cast_util.wrap_str, '射源种类', column_type=str),
                "J": ExcelFiledProperty('J', cast_util.wrap_str, '设备型号射源种类', column_type=str),
                "K": ExcelFiledProperty('K', cast_util.wrap_str, '焦点尺寸', column_type=str),
                "L": ExcelFiledProperty('L', cast_util.wrap_str, '增感方式', column_type=str),
                "M": ExcelFiledProperty('M', cast_util.wrap_str, '胶片牌号', column_type=str),
            }
            guide_data_list = excel_util.read_data_list(file_path=context.input_file_path, sheet=1,
                                                        filed_mapping=guide_filed_mapping)
            guide_data_map = {}
            for guide_d in guide_data_list:
                guide_specification = guide_d['A']
                if guide_specification is None:
                    guide_specification = ''

                guide_specification = guide_specification.replace('×', '*')
                guide_specification = guide_specification.replace('Φ', 'φ')

                guide_data_map[guide_specification] = guide_d

        data_list_map: Dict[str, List[dict]] = {}
        for data in data_list:
            order_id = data.get('orderId', None)
            if order_id == '' or order_id is None:
                continue
            if order_id not in data_list_map:
                data_list_map[order_id] = []
            data_list_map[order_id].append(data)

        for order_id, order_data_list in data_list_map.items():
            temp_list = []
            valid = True
            for data in order_data_list:
                expect_date_str = map_util.get(context.global_param, 'expectDate')
                expect_date = awe_date_util.parse_dot_date_time(expect_date_str)

                date_str = map_util.get(data, 'completeDate')
                date = awe_date_util.parse_dot_date_time(date_str)

                if (date is None and expect_date is None) or date > expect_date:
                    valid = False
                    break

                if data.get('isOk') == '合格' or data.get('isOk') == '不合格':
                    check_count = data.get('checkCount', 0)
                    if check_count is None or check_count <= 0:
                        continue
                    temp_list.append(data)

                    specification = data.get('specification', None)
                    if specification is None:
                        specification = ''

                    specification = specification.replace('×', '*')
                    specification = specification.replace('Φ', 'φ')

                    data['_guide.A'] = guide_data_map.get(specification, {}).get('A', None)
                    data['_guide.B'] = guide_data_map.get(specification, {}).get('B', None)
                    data['_guide.C'] = guide_data_map.get(specification, {}).get('C', None)
                    data['_guide.D'] = guide_data_map.get(specification, {}).get('D', None)
                    data['_guide.E'] = guide_data_map.get(specification, {}).get('E', None)
                    data['_guide.F'] = guide_data_map.get(specification, {}).get('F', None)
                    data['_guide.G'] = guide_data_map.get(specification, {}).get('G', None)
                    data['_guide.H'] = guide_data_map.get(specification, {}).get('H', None)
                    data['_guide.I'] = guide_data_map.get(specification, {}).get('I', None)
                    data['_guide.J'] = guide_data_map.get(specification, {}).get('J', None)
                    data['_guide.K'] = guide_data_map.get(specification, {}).get('K', None)
                    data['_guide.L'] = guide_data_map.get(specification, {}).get('L', None)
                    data['_guide.M'] = guide_data_map.get(specification, {}).get('M', None)

                    if check_count == 6:
                        data['_pieceNo'] = '1-2'
                        for i in range(2, 7):
                            if i == 6:
                                new_data = {'_pieceNo': '6-1'}
                            else:
                                new_data = {'_pieceNo': f'{i}-{(i + 1)}'}
                            temp_list.append(new_data)
                            new_data['orderId'] = data.get('orderId', None)
                            new_data['sampleNo'] = data.get('sampleNo', None)
                            new_data['kindNo'] = data.get('kindNo', None)
                            new_data['empId'] = data.get('empId', None)
                            new_data['specification'] = data.get('specification', None)
                    else:
                        data['_pieceNo'] = '1'
                        for i in range(0, check_count - 1):
                            new_data = {'_pieceNo': f'{i + 2}', 'orderId': data.get('orderId', None),
                                        'sampleNo': data.get('sampleNo', None), 'kindNo': data.get('kindNo', None),
                                        'empId': data.get('empId', None),
                                        'specification': data.get('specification', None)}
                            temp_list.append(new_data)
                else:
                    valid = False
                    break
            if valid and len(temp_list) > 0:
                for _ in temp_list:
                    new_data_list.append(_)
        return new_data_list

    report_generator = MergedExcelReportGenerator(template_path=None, filed_mapping=filed_mapping,
                                                  divide_key=divide_key,
                                                  doc_global_data_param_config_list=doc_global_data_param_config_list,
                                                  excel_header_value_setter_list=excel_header_value_setter_list,
                                                  excel_columns_value_setter_list=excel_columns_value_setter_list,
                                                  data_preparer=data_preparer,
                                                  merge_fun_dict={'completeDate': date_merge_fun}
                                                  )
    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='质量评定台账Excel生成', processor=report_generator, biz_code='rt_excel'))
    base_ui_config = BaseUIConfig('质量评定台账Excel生成', processor_widget)

    return base_ui_config


# 20240310 射线检测记录
def build_ray_dect_record_20240310_config():

    # 在外输入
    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('customorCompany', None, '委托单位', False),
    ]

    company = doc_global_data_param_config_list[1].title

    divide_key = 'orderId'
    # 从表格中读取的数据
    filed_mapping = {

        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/㎡/点)'),

    }

    cell_resource_list = [

        # 承包单位-委托单位
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='customorCompany',
                                               style=DocCellStyle(font_cn='楷体')),
        # 委托单编号
        DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=3, mapping_key=divide_key,
                                            style=DocCellStyle(font_cn='楷体', font_size=10)),
        # 焊接方法
        DataListCellParagraphAddRunResource(table_index=0, row=4, column_view_index=1, mapping_key='weldMethod',
                                            style=DocCellStyle(font_cn='楷体')),
        # 检测比例
        DataListCellParagraphAddRunResource(table_index=0, row=6, column_view_index=3, mapping_key='checkRatioKind',
                                            style=DocCellStyle(font_cn='楷体')),
        # 合格级别
        DataListCellParagraphAddRunResource(table_index=0, row=6, column_view_index=7, mapping_key='level',
                                            style=DocCellStyle(font_cn='楷体')),


        # # 洗片人时间
        # # paragraph_index 同一格中行数
        # DataListCellParagraphAddRunResource(table_index=0, row=35, column_view_index=0,
        #                                     style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=2,
        #                                     cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='completeDate',
        #                                     data_list_sort_func=awe_date_util.parse_dot_date_time,
        #                                     data_list_sort_reverse=True),
        # # 拍片人时间
        # DataListCellParagraphAddRunResource(table_index=0, row=35, column_view_index=1,
        #                                     style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=2,
        #                                     cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='completeDate',
        #                                     data_list_sort_func=awe_date_util.parse_dot_date_time,
        #                                     data_list_sort_reverse=True),
        # # 审核人时间
        # DataListCellParagraphAddRunResource(table_index=0, row=35, column_view_index=2,
        #                                     style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=2,
        #                                     cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='completeDate',
        #                                     data_list_sort_func=awe_date_util.parse_dot_date_time,
        #                                     data_list_sort_reverse=True),

    ]

    column_cell_resource_list = [
        # 检件编号
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=18, column_view_index=1,
                                                  mapping_data_key='sampleNo',style=DocCellStyle(font_cn='楷体')),

        # 焊口编号
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=18, column_view_index=2,
                                                  mapping_data_key='kindNo', style=DocCellStyle(font_cn='楷体')),
        # 焊工编号
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=18, column_view_index=3,
                                                  mapping_data_key='empId', style=DocCellStyle(font_cn='楷体')),

        # 备注-完成日期
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=18, column_view_index=5,
                                                  mapping_data_key='completeDate', style=DocCellStyle(font_cn='楷体')),
    ]

    template_path = os.path.join(os.getcwd(), 'template/ray_dect_record_20240310/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RayAdditionalHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list,
                                               merge_fun_dict={'completeDate': date_merge_fun})

    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='射线检测记录', processor=report_generator, biz_code='ray_dect_record_20240310'))

    base_ui_config = BaseUIConfig('射线检测记录', processor_widget)

    return base_ui_config


# 20240310 射线检测记录（续）
def build_ray_dect_record_con_20240310_config():

    divide_key = 'orderId'
    filed_mapping = {

        # 从表格中读取的数据
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号', column_type=str),
        "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号', column_type=str),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_percent, '检测比例'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/㎡/点)'),

        "black": ExcelFiledProperty('W', cast_util.wrap_str, '底片黑度'),

    }

    column_cell_resource_list = [

        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=0,
                                                  mapping_data_key='sampleNo',
                                                  style=DocCellStyle(font_cn='楷体',font_size=9)),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=1,
                                                  mapping_data_key='kindNo',
                                                  style=DocCellStyle(font_cn='楷体',font_size=9)),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=3,
                                                  mapping_data_key='empId',
                                                  style=DocCellStyle(font_cn='楷体',font_size=9)),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=4,
                                                  mapping_data_key='specification',
                                                  style=DocCellStyle(font_cn='楷体', font_size=8)),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=5,
                                                  mapping_data_key='black',
                                                  style=DocCellStyle(font_cn='楷体', font_size=9)),
        # 备注-完成日期
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=4, column_view_index=11,
                                                  mapping_data_key='completeDate',
                                                  style=DocCellStyle(font_cn='楷体',font_size=9)),


    ]
    # 界面上需要输入的
    doc_global_data_param_config_list = [
        DocGlobalParamConfig('projectName', None, '工程名称', False),
        DocGlobalParamConfig('customorCompany', None, '委托单位', False),
    ]

    cell_resource_list = [

        # 委托单编号
        DataListCellParagraphAddRunResource(table_index=0, row=0, column_view_index=3, mapping_key=divide_key,
                                            style=DocCellStyle(font_cn='楷体',font_size=10)),


        # 检测比例
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='checkRatioKind',
                                            style=DocCellStyle(font_cn='楷体',font_size=10)),

        # 合格级别
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3, mapping_key='level',
                                            style=DocCellStyle(font_cn='楷体',font_size=10)),

        # # 评片人时间
        # DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=0,
        #                                     style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=4,
        #                                     cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,mapping_key='completeDate',
        #                                     data_list_sort_func=awe_date_util.parse_dot_date_time,
        #                                     data_list_sort_reverse=True),
        #
        # # 审核人时间
        # DataListCellParagraphAddRunResource(table_index=0, row=23, column_view_index=1,
        #                                     style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=4,
        #                                     cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,mapping_key='completeDate',
        #                                     data_list_sort_func=awe_date_util.parse_dot_date_time,
        #                                     data_list_sort_reverse=True),

    ]

    template_path = os.path.join(os.getcwd(), 'template/ray_dect_record_con_20240310/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RayAdditionalHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               # data_preparer=None,
                                               # data_preparer=data_preparer_con,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list,
                                               merge_fun_dict={'completeDate': date_merge_fun})

    processor_widget = ProcessorQWidget(
        ProcessorUIParam(title='射线检测记录续', processor=report_generator, biz_code='ray_dect_record_con_20240310'))

    base_ui_config = BaseUIConfig('射线检测记录续', processor_widget)

    return base_ui_config


if __name__ == '__main__':
    # Must run before QWidgets init.
    print(os.getcwd())
    app = QApplication(sys.argv)

    ui_config_list = [
        build_ray_config(),
        build_rt_config(),
        build_surface_config(),
        build_record_excel_config(),
        build_ray_dect_record_20240310_config(),
        build_ray_dect_record_con_20240310_config(),

    ]
    ui = HomePageQWidget(ui_config_list)
    ui.show()

    sys.exit(app.exec_())
