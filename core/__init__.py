import os.path
from abc import ABCMeta, abstractmethod

import docx
import pandas
from typing import List, Dict

from docx import Document

from .util import excelTitleToIndex

DEFAULT_SHEET_NAME = 'Sheet1'


class FiledProperty():
    """
    value_cast is a function: val = value_cast(value)
    """
    column_title: str
    value_cast= None


class AbcReportGenerator(metaclass=ABCMeta):
    """
    Abstract excel to words report generator
    Input: one excel file
    Output: multi word files
    """

    template_path = None
    filed_mapping: Dict[str, FiledProperty] = None
    divide_key = None

    # key -> cast
    #

    # args: number([0:N)), totalCount(N), success:bool, exception: BaseException
    on_finsh_one = None

    def __init__(self, template_path: str, on_finish_one):
        self.template_path = template_path
        self.on_finsh_one = on_finish_one

    def execute(self, file_path: str, target_dir: str, sheet: str = DEFAULT_SHEET_NAME, global_data: dict = None):
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
                doc = self.process(data_list=sub_data_list, template_doc=template_doc, global_data=global_data)
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

            for k, filed_property in self.filed_mapping:
                filed_property: FiledProperty
                col_idx = excelTitleToIndex(filed_property.column_title)
                val = self._castValue(row[col_idx], filed_property.value_cast)
                raw_data[k] = val

            data_list.append(raw_data)

        return data_list

    @abstractmethod
    def process(self, data_list: list, template_doc: Document, global_data: dict) -> Document:
        pass

    def get_file_name(self, key):
        return f'{key}.docx'

    def _castValue(self, value, cast):
        if cast is None:
            return value
        return cast(value)
