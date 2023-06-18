import os.path
import traceback
from abc import ABCMeta, abstractmethod

import docx
import pandas
from typing import List, Dict

from docx import Document

from . import DEFAULT_SHEET_NAME
from .util import excel_title_to_index


class ExcelFiledProperty:
    """
    value_cast is a function: val = value_cast(value)
    """

    column_title: str
    value_cast = None
    desc: str

    def __init__(self, column_title, value_cast, desc: str) -> None:
        super().__init__()
        self.column_title = column_title
        self.value_cast = value_cast
        self.desc = desc


class DocGlobalParamConfig:
    def __init__(self, filed: str, default_value=None, title: str = None, required: bool = False):
        self.filed = filed
        self.title = title
        self.required = required
        self.default_value = default_value


def default_on_finish_one(number: int, total_cnt: int, success: bool, exception):
    pass


class ReportGenerator(metaclass=ABCMeta):
    """
    Abstract excel to words report generator
    Input: one excel file
    Output: multi word files
    """

    # key -> cast
    #

    # args: number([0:N)), totalCount(N), success:bool, exception: BaseException

    def __init__(self, template_path: str, filed_mapping: Dict[str, ExcelFiledProperty], divide_key: str,
                 doc_global_data_param_config_list: List[DocGlobalParamConfig]=None):
        """
        :param template_path: path to template docx file
        :param filed_mapping:
        :param divide_key: unique key
        """
        super().__init__()
        self.template_path = template_path
        self.filed_mapping = filed_mapping
        self.divide_key = divide_key
        self.doc_global_data_param_config_list = doc_global_data_param_config_list

    def execute(self, file_path: str, target_dir: str, sheet: str = DEFAULT_SHEET_NAME, global_param: dict = None, on_finish_one=default_on_finish_one):
        """
        :param on_finish_one: callback, on_finish_one(number: int, total_cnt: int, success: bool, exception: BaseException)
        """
        if global_param is None:
            global_param = {}
        self._check_and_set_default_global_param(global_param)
        data_list = self.__read(file_path, sheet)
        raw_data_map = {}
        for raw_data in data_list:
            u_val = raw_data[self.divide_key]
            if u_val is None:
                continue
            if u_val not in raw_data_map:
                raw_data_map[u_val] = []
            raw_data_map[u_val].append(raw_data)

        p = 1

        callback = on_finish_one

        if callback is None:
            callback = default_on_finish_one

        for key, sub_data_list in raw_data_map.items():
            try:
                template_doc = docx.Document(self.template_path)
                doc = self._process(data_list=sub_data_list, template_doc=template_doc, global_param=global_param)
                file_name = self._get_file_name(key)
                save_path = os.path.join(target_dir, file_name)
                self._save(doc, file_path=save_path)
                callback(p, len(raw_data_map.keys()), True, None)
            except BaseException as e:
                traceback.print_exc()
                callback(p, len(raw_data_map.keys()), False, e)
            finally:
                p = p + 1

    def __read(self, file_path: str, sheet: str) -> List[dict]:
        raw_datas = pandas.read_excel(file_path, sheet)
        data_list = []
        for i in range(raw_datas.shape[0]):
            row = raw_datas.iloc[i]
            raw_data = {}

            for k, filed_property in self.filed_mapping.items():
                filed_property: ExcelFiledProperty
                col_idx = excel_title_to_index(filed_property.column_title)
                val = self._castValue(row[col_idx], filed_property.value_cast)
                raw_data[k] = val

            data_list.append(raw_data)

        return data_list

    def _check_and_set_default_global_param(self, global_param: dict):
        if self.doc_global_data_param_config_list is None:
            return
        for config in self.doc_global_data_param_config_list:
            config: DocGlobalParamConfig
            if config.required and config.filed not in global_param:
                raise ValueError(f'缺少参数:{config.title}')

    @abstractmethod
    def _process(self, data_list: list, template_doc: Document, global_param: dict) -> Document:
        pass

    def _get_file_name(self, key):
        return f'{key}.docx'

    def _castValue(self, value, cast):
        if cast is None:
            return value
        return cast(value)

    def _save(self, doc: Document, file_path: str):
        doc.save(file_path)
