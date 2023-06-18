from abc import ABCMeta, abstractmethod

from docx import Document
from typing import Dict, List

from docx.enum.table import WD_TABLE_ALIGNMENT

from .base import ReportGenerator, ExcelFiledProperty
from .style import DocCellStyle



class HeaderResource(metaclass=ABCMeta):
    """
    用于处理Word表格头数据
    """

    @abstractmethod
    def set(self, doc: Document, data_list: List[Dict], global_data: dict):
        pass


class ColumnCellsResource(metaclass=ABCMeta):
    """
    用于处理word内表格内某一列数据
    """

    def __init__(self, table_index: int, table_data_start_row, column: int, style: DocCellStyle = DocCellStyle()):
        self.table_index = table_index
        self.table_data_start_row = table_data_start_row
        self.column = column
        self.style = style

    @abstractmethod
    def get_value(self, data: dict, global_data: dict):
        pass

    def set(self, cur_index: int, doc: Document, data: dict, global_data: dict):
        val = self.get_value(data, global_data)
        if val is None:
            return
        table = doc.tables[self.table_index]
        row = table.rows[self.table_data_start_row + cur_index]
        cell = row.cells[self.column]
        cell.text = val
        if self.style is not None:
            cell.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER


class CellResource(metaclass=ABCMeta):
    """
    用于处理表格内某一个cell数据
    """

    def __init__(self, table_index: int, row: int, column: int):
        self.table_index = table_index
        self.row = row
        self.column = column

    @abstractmethod
    def get_value(self, data_list: List[dict], global_data: dict):
        pass

    def set(self, doc: Document, data_list: List[dict], global_data: dict):
        val = self.get_value(data_list, global_data)
        if val is None:
            return
        table = doc.tables[self.table_index]
        cell = table.cell(self.row, self.column)
        cell.text = val


class MappingColumnCellsResource(ColumnCellsResource):

    def __init__(self, table_index: int, table_data_start_row, column: int, mapping_data_key: str):
        super().__init__(table_index, table_data_start_row, column)
        self.mapping_data_key = mapping_data_key

    def get_value(self, data: dict, global_data: dict):
        if self.mapping_data_key in data:
            return data[self.mapping_data_key]
        return None


class SimpleTableDocGenerator(ReportGenerator):

    def __init__(self, template_path: str, filed_mapping: Dict[str, ExcelFiledProperty], divide_key: str, on_finish_one,
                 header_resource: HeaderResource, column_cell_resource_list: List[ColumnCellsResource],
                 cell_resource_list: List[CellResource]):
        super().__init__(template_path, filed_mapping, divide_key, on_finish_one)
        self.header_resource = header_resource
        self.column_cell_resource_list = column_cell_resource_list
        self.cell_resource_list = cell_resource_list

    ##### Template descript ####
    # excel key to word table column index
    table_mapping: Dict[str, int]

    # start of data list
    data_start_row_of_table: int

    def _process(self, data_list: list, template_doc: Document, global_data: dict) -> Document:
        if self.header_resource is not None:
            self.header_resource.set(doc=template_doc, data_list=data_list, global_data=global_data)
        if self.column_cell_resource_list is not None:
            for i in range(len(data_list)):
                data = data_list[i]
                for col_res in self.column_cell_resource_list:
                    col_res.set(i, template_doc, data, global_data)
        if self.cell_resource_list is not None:
            for cell_res in self.cell_resource_list:
                cell_res.set(template_doc, data_list, global_data)

        return template_doc
