from abc import ABC
from typing import List, Dict

from docx import Document

from awe_report_generator.core.simple_table_doc_generator import HeaderResource, CellResource


class RTHeaderResource(HeaderResource, ABC):

    def set(self, doc: Document, data_list: List[Dict], global_data: dict):
        project_name = global_data['projectName']
        doc.paragraphs[1].add_run(project_name)
        doc.paragraphs[2].add_run(self.get_one_value(data_list, 'unitName'))


class RTSummaryCellResource(CellResource):
    SUMMARY_FORMAT_ONE = '说明：共检测{data_size}道,合格{ok_data_size}道，不合格{bad_data_size}道，其中返修{bad_check_count}张，共计{check_count}张。'
    SUMMARY_FORMAT_TWO = '其中γ射线{ray}张。'

    def get_value(self, data_list: List[dict], global_data: dict):
        if data_list is None:
            return None
        data_size = len(data_list)
        ok_data_size = 0

        check_count = 0
        ok_check_count = 0
        ray = 0

        for data in data_list:
            if data.get('isOk', '') != '合格':
                pass
            else:
                ok_data_size += 1
            check_count += data.get('checkCount', 0)
            ok_check_count += data.get('okCount', 0)
            ray += data.get('ray', 0)

        val = self.SUMMARY_FORMAT_ONE.format(data_size=data_size, ok_data_size=ok_data_size,
                                             bad_data_size=data_size - ok_data_size, bad_check_count=check_count - ok_check_count, check_count=check_count)

        if ray > 0:
            val = val + self.SUMMARY_FORMAT_TWO.format(ray=ray)

        return val
