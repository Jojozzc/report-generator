import os
import sys

from PyQt5.QtWidgets import QApplication
from docx.enum.table import WD_TABLE_ALIGNMENT

from awe_report_generator.biz.rt.base import RTHeaderResource, RTSummaryCellResource2
from awe_report_generator.biz.rt.base import RTSummaryCellResource
from awe_report_generator.core.base import ExcelFiledProperty, DocGlobalParamConfig
from awe_report_generator.core.simple_table_doc_generator import (
    SimpleTableDocGenerator, MappingColumnCellsParagraphAddRunResource,
    SimpleCalculationColumnCellsParagraphAddRunResource,
    HeaderResource,
    GlobalParamCellParagraphAddRunResource,
    DataListCellParagraphAddRunResource, MergedDataCellParagraphAddRunResource
)
from awe_report_generator.core.style import DocCellStyle
from awe_report_generator.core.util import cast_util, date_util as awe_date_util
from awe_report_generator.ui import BaseUIConfig
from awe_report_generator.ui.base import HomePageQWidget
from awe_report_generator.ui.processor_ui import ProcessorQWidget, ProcessorUIParam
from awe_report_generator.core.util import array_util


def date_merge_fun(date_list):
    if date_list is None or len(date_list) == 0:
        return None
    date_list = array_util.sort_date_array(date_list, True)
    return date_list[0]


def build_rt_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_int_str, '焊工号'),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
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
        MergedDataCellParagraphAddRunResource(table_index=0, row=1, column_view_index=3, mapping_key='completeDate',
                                            style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='detectionMethod',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3,
                                               mapping_key='testingStandards', style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=5, mapping_key='level',
                                            style=DocCellStyle(font_cn='楷体')),
        RTSummaryCellResource(table_index=0, row=21, column_view_index=0,
                              style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.LEFT)),
        MergedDataCellParagraphAddRunResource(table_index=0, row=22, column_view_index=3,
                                            style=DocCellStyle(alignment=WD_TABLE_ALIGNMENT.RIGHT), paragraph_index=3,
                                            run_index=None, cast_value_to_str=awe_date_util.get_YYYYmmdd_cn,
                                            mapping_key='completeDate'),
    ]
    template_path = os.path.join(os.getcwd(), 'template/rt/TEMPLATE.docx')
    report_generator = SimpleTableDocGenerator(template_path=template_path, filed_mapping=filed_mapping,
                                               divide_key=divide_key,
                                               header_resource=RTHeaderResource(),
                                               column_cell_resource_list=column_cell_resource_list,
                                               cell_resource_list=cell_resource_list,
                                               doc_global_data_param_config_list=doc_global_data_param_config_list, merge_fun_dict={'completeDate' : date_merge_fun})
    processor_widget = ProcessorQWidget(ProcessorUIParam(title='RT结果通知单台账', processor=report_generator))
    base_ui_config = BaseUIConfig('RT结果通知单台账', processor_widget)

    return base_ui_config


def build_ray_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_int_str, '焊工号'),
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
                                                  mapping_data_key='specification', style=DocCellStyle(font_cn='楷体', font_size=7.5)),
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
                                            cast_value_to_str=awe_date_util.get_YYYYmmdd_cn, mapping_key='orderDate'),
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
                                               doc_global_data_param_config_list=doc_global_data_param_config_list, merge_fun_dict={'completeDate' : date_merge_fun})

    processor_widget = ProcessorQWidget(ProcessorUIParam(title='射线检测委托台账', processor=report_generator))

    base_ui_config = BaseUIConfig('射线检测委托台账', processor_widget)

    return base_ui_config


def build_surface_config():
    divide_key = 'orderId'
    filed_mapping = {
        "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
        "completeDate": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
        divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
        "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
        "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
        "empId": ExcelFiledProperty('F', cast_util.wrap_int_str, '焊工号'),
        "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格(mm)'),
        "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
        "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
        "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
        "isOk": ExcelFiledProperty('K', cast_util.wrap_str, '返修补片'),
        "baseSpecificationAndCnt": ExcelFiledProperty('L', cast_util.wrap_str, '底片规格/张'),
        "unitName": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
        "checkCount": ExcelFiledProperty('M', cast_util.wrap_int, '张数'),
        "okCount": ExcelFiledProperty('N', cast_util.wrap_int, '合格数量'),
        "ray": ExcelFiledProperty('P', cast_util.wrap_str, 'γ射线'),
        "detectionCount": ExcelFiledProperty('U', cast_util.wrap_str, '检测数量(道/m/m2/点)'),

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
                                                  mapping_data_key='detectionCount',
                                                  style=DocCellStyle(font_cn='楷体')),
        MappingColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=5,
                                                  mapping_data_key='okCount', style=DocCellStyle(font_cn='楷体')),
        SimpleCalculationColumnCellsParagraphAddRunResource(table_index=0, table_data_start_row=5, column_view_index=6,
                                                            mapping_data_key_1='checkCount',
                                                            mapping_data_key_2='okCount',
                                                            operation='-', style=DocCellStyle(font_cn='楷体')),
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
                                            mapping_key='completeDate', style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=1, column_view_index=1, mapping_key='customerCompany',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=1, mapping_key='detectionMethod',
                                               style=DocCellStyle(font_cn='楷体')),
        GlobalParamCellParagraphAddRunResource(table_index=0, row=2, column_view_index=3,
                                               mapping_key='testingStandards', style=DocCellStyle(font_cn='楷体')),
        DataListCellParagraphAddRunResource(table_index=0, row=2, column_view_index=5, paragraph_index=0,
                                            mapping_key='level', style=DocCellStyle(font_cn='楷体')),
        MergedDataCellParagraphAddRunResource(table_index=0, row=25, column_view_index=3,
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
                                               doc_global_data_param_config_list=doc_global_data_param_config_list, merge_fun_dict={'completeDate' : date_merge_fun})
    processor_widget = ProcessorQWidget(ProcessorUIParam(title='表面结果通知单台账', processor=report_generator))
    base_ui_config = BaseUIConfig('表面结果通知单台账', processor_widget)

    return base_ui_config


if __name__ == '__main__':
    # Must run before QWidgets init.
    print(os.getcwd())
    app = QApplication(sys.argv)

    ui_config_list = [
        build_ray_config(),
        build_rt_config(),
        build_surface_config(),
    ]
    ui = HomePageQWidget(ui_config_list)
    ui.show()

    sys.exit(app.exec_())
