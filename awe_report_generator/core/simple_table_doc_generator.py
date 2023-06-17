from docx import Document
from typing import Dict, List

from .base import ReportGenerator


def default_mapping_handler(data_list: List[dict], global_data):
    pass


class GlobalHeaderMapping():
    filed: str
    desc: str

    # mapping_handler is a function: mapping_handler()
    mapping_handler = None


class SimpleTableDocGenerator(ReportGenerator):
    ##### Template descript ####

    # excel key to word table column index
    table_mapping: Dict[str, int]

    # start of data list
    data_start_row_of_table: int

    def _process(self, data_list: list, template_doc: Document, global_data: dict) -> Document:
        pass
