from docx import Document

from awe_report_generator.core.base import (
    ReportGenerator,
    ExcelFiledProperty,
)
from awe_report_generator.core.util import cast_util


class RadioCheckGenerator(ReportGenerator):
    def __init__(self, template_path: str, on_finish_one):
        divide_key = 'orderId'
        filed_mapping = {
            "orderDate": ExcelFiledProperty('A', cast_util.wrap_str, '委托日期'),
            "complete_date": ExcelFiledProperty('B', cast_util.wrap_str, '完成时间'),
            divide_key: ExcelFiledProperty('C', cast_util.wrap_str, '委托单编号'),
            "sampleNo": ExcelFiledProperty('D', cast_util.wrap_str, '检件编号'),
            "kindNo": ExcelFiledProperty('E', cast_util.wrap_str, '焊口编号'),
            "empId": ExcelFiledProperty('F', cast_util.wrap_str, '焊工号'),
            "specification": ExcelFiledProperty('G', cast_util.wrap_str, '规格（mm）'),
            "material": ExcelFiledProperty('H', cast_util.wrap_str, '材质'),
            "level": ExcelFiledProperty('I', cast_util.wrap_str, '合格级别'),
            "checkRatioKind": ExcelFiledProperty('J', cast_util.wrap_str, '检测比列'),
            "projectCompany": ExcelFiledProperty('Q', cast_util.wrap_str, '单元名称'),
            "method": ExcelFiledProperty('R', cast_util.wrap_str, '焊接方法'),
            "lineNo": ExcelFiledProperty('T', cast_util.wrap_str, '单线号'),
        }
        super().__init__(template_path, filed_mapping, divide_key, on_finish_one)


    def _process(self, data_list: list, template_doc: Document, global_param: dict) -> Document:
        """
        :param global_param:
        :return:
        """
        if global_param is None:
            raise ValueError('global_data is None')
        table = template_doc.tables[0]

        return template_doc

    def _save(self, doc: Document, file_path: str):
        print(f'save to {file_path}')




