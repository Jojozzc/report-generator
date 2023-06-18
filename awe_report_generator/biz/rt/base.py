from abc import ABC
from typing import List, Dict

from docx import Document

from awe_report_generator.core.simple_table_doc_generator import HeaderResource


class RTHeaderResource(HeaderResource, ABC):
    def set(self, doc: Document, data_list: List[Dict], global_data: dict):
        pass