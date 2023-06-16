import os.path
from abc import ABCMeta, abstractmethod

import docx
import pandas
from typing import List, Dict

from docx import Document

"""
    Abstract report generator.
"""

DEFAULT_SHEET_NAME = 'Sheet1'


class AbcReportGenerator(metaclass=ABCMeta):
    template_path = None
    key_col_mapping: dict = None
    divide_key = None

    # key -> cast
    # cast is a function val = convert(value)
    key_value_cast: dict = None

    # args: number([0:N)), totalCount(N), success:bool, exception: BaseException
    on_finsh_one = None

    def execute(self, file_path: str, target_dir: str, sheet: str = DEFAULT_SHEET_NAME):
        data_list = self.read(file_path, sheet)
        raw_data_map = {}
        for raw_data in data_list:
            u_val = raw_data[self.divide_key]
            if u_val is None:
                continue
            if u_val not in raw_data_map:
                raw_data_map[u_val] = []
            raw_data_map[u_val].append(raw_data)

        p = 0

        callback = self.on_finsh_one

        if callback is None:
            callback = lambda number, total_cnt, success, exception: None

        for key, sub_data_list in raw_data_map:
            try:
                template_doc = docx.Document(self.template_path)
                doc = self.process(data_list=sub_data_list, template_doc=template_doc)
                file_name = self.get_file_name(key)
                save_path = os.path.join(target_dir, file_name)
                doc.save(save_path)
                callback(p, len(raw_data_map.keys()), True, None)
            except BaseException as e:
                callback(p, len(raw_data_map.keys()), False, e)
            finally:
                p = p + 1

    def read(self, file_path: str, sheet: str) -> List[dict]:
        raw_datas = pandas.read_excel(file_path, sheet)
        data_list = []
        for i in range(raw_datas.shape[0]):
            row = raw_datas.iloc[i]
            raw_data = {}

            for k, col_idx in self.key_col_mapping:
                val = self._castValue(row[col_idx], self.key_value_cast[k])
                raw_data[k] = val

            data_list.append(raw_data)

        return data_list

    @abstractmethod
    def process(self, data_list: list, template_doc: Document) -> Document:
        pass

    def get_file_name(self, key):
        return f'{key}.docx'

    def _castValue(self, value, cast):
        if cast is None:
            return value
        return cast(value)
