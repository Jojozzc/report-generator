from abc import ABCMeta
from typing import Dict, List

from docx import Document
from xlsxwriter.worksheet import Worksheet

from awe_report_generator.core.base import ReportGenerator, ExcelFiledProperty, DivideMode, OutputFileMode, \
    DocGlobalParamConfig, ValueGetter


class ExcelColumnsValueSetter(metaclass=ABCMeta):
    def __init__(self, column: str, value_getter: ValueGetter):
        self.column = column
        self.value_getter = value_getter

    def set(self, worksheet: Worksheet, row: int, data: dict, data_list: List[dict], global_data: dict, merged_data: dict):
        val = self.value_getter.get_value(data=data, data_list=data_list, global_data=global_data, merged_data=merged_data)
        if val is not None:
            worksheet.write(f'{self.column}{row}', val)


class MergedExcelReportGenerator(ReportGenerator):
    def __init__(self, template_path: str, filed_mapping: Dict[str, ExcelFiledProperty], divide_key: str,
                 doc_global_data_param_config_list: List[DocGlobalParamConfig] = None,
                 merge_fun_dict: dict = None,
                 excel_header_value_setter_list:List[ExcelColumnsValueSetter]=None,
                 excel_columns_value_setter_list:List[ExcelColumnsValueSetter] = None,
                 data_preparer=None,
                 ):
        super().__init__(template_path=template_path, filed_mapping=filed_mapping, divide_key=divide_key,
                         divide_mode=DivideMode.DIVIDE_MODE_NO_DIVIDE,
                         output_file_mode=OutputFileMode.EXCEL,
                         doc_global_data_param_config_list=doc_global_data_param_config_list,
                         merge_fun_dict=merge_fun_dict, data_preparer=data_preparer)
        self.excel_header_value_setter_list = excel_header_value_setter_list
        self.excel_columns_value_setter_list = excel_columns_value_setter_list

    def _process_word(self, data_list: list, template_doc: Document, global_param: dict, merged_data: dict) -> Document:
        return super()._process_word(data_list=data_list, template_doc=template_doc, global_param=global_param,
                                     merged_data=merged_data)

    def _process_excel(self, data_list: list, worksheet: Worksheet, global_param: dict, merged_data: dict):
        if self.excel_header_value_setter_list is not None:
            for value_setter in self.excel_header_value_setter_list:
                value_setter.set(worksheet=worksheet, row=1, data=None, data_list=data_list,
                                 global_data=global_param, merged_data=merged_data)

        if self.excel_columns_value_setter_list is not None:
            for value_setter in self.excel_columns_value_setter_list:
                for i in range(len(data_list)):
                    data = data_list[i]
                    value_setter.set(worksheet=worksheet, row=i + 2, data=data, data_list=data_list, global_data=global_param, merged_data=merged_data)

