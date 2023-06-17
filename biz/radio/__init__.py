from docx import Document

from core import (
    ReportGenerator,
    FiledProperty,
)
from core.util import cast_util


class RadioCheckGenerator(ReportGenerator):
    def __init__(self, template_path: str, on_finish_one):
        super().__init__(template_path, on_finish_one)
        self.divide_key = 'orderId'
        self.filed_mapping = {
            "orderDate": FiledProperty('A', cast_util.wrap_str, '委托日期'),
            "complete_date": FiledProperty('B', cast_util.wrap_str, '完成时间'),
            "orderId": FiledProperty('C', cast_util.wrap_str, '委托单编号'),
            "sampleNo": FiledProperty('D', cast_util.wrap_str, '检件编号'),
            "kindNo": FiledProperty('E', cast_util.wrap_str, '焊口编号'),
            "empId": FiledProperty('F', cast_util.wrap_str, '焊工号'),
            "specification": FiledProperty('G', cast_util.wrap_str, '规格（mm）'),
            "material": FiledProperty('H', cast_util.wrap_str, '材质'),
            "level": FiledProperty('I', cast_util.wrap_str, '合格级别'),
            "checkRatioKind": FiledProperty('J', cast_util.wrap_str, '检测比列'),
            "projectCompany": FiledProperty('Q', cast_util.wrap_str, '单元名称'),
            "method": FiledProperty('R', cast_util.wrap_str, '焊接方法'),
            "lineNo": FiledProperty('T', cast_util.wrap_str, '单线号'),
        }

    def _process(self, data_list: list, template_doc: Document, global_data: dict) -> Document:
        return template_doc

    def _save(self, doc: Document, file_path: str):
        print(f'save to {file_path}')




