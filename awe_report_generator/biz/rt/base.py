from abc import ABC
from typing import List, Dict

from docx import Document

from awe_report_generator.core.simple_table_doc_generator import HeaderResource, CellResource


class RTHeaderResource(HeaderResource, ABC):

    def set(self, doc: Document, data_list: List[Dict], global_data: dict):
        project_name = global_data['projectName']
        doc.paragraphs[1].add_run(project_name)
        doc.paragraphs[2].add_run(self.get_one_value(data_list, 'unitName'))


class DescCellResource(CellResource):
    def get_value(self, data_list: List[dict], global_data: dict):
        pass