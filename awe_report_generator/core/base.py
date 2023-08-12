import os.path
import traceback
from abc import ABCMeta, abstractmethod

import docx
import pandas
from typing import List, Dict

from docx import Document

from . import DEFAULT_SHEET
from .util import excel_title_to_index


class ExcelFiledProperty:
    """
    value_cast is a function: val = value_cast(value)
    """

    def __init__(self, column_title, value_cast, desc: str, required=False, column_type=None, valid_check=None) -> None:
        '''
        :param column_type: see dtype in https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
        :param valid_check: function valid(value_cast(value))
        '''
        super().__init__()
        self.column_title = column_title
        self.value_cast = value_cast
        self.desc = desc
        self.required = required
        self.column_type = column_type
        self.valid_check = valid_check


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
                 doc_global_data_param_config_list: List[DocGlobalParamConfig] = None,
                 merge_fun_dict: dict = None):
        """
        :param template_path: path to template docx file
        :param filed_mapping:
        :param divide_key: unique key

        dtype: Data type for data or columns. E.g. {‘a’: np.float64, ‘b’: np.int32} Use object to preserve data as stored in Excel and not interpret dtype. If converters are specified, they will be applied INSTEAD of dtype conversion.
            https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
        """
        super().__init__()
        self.template_path = template_path
        self.filed_mapping = filed_mapping
        self.divide_key = divide_key
        self.doc_global_data_param_config_list = doc_global_data_param_config_list
        self.merge_fun_dict = merge_fun_dict
        self.dtype = {}
        for key, property in filed_mapping.items():
            if property is not None and property.column_type is not None:
                self.dtype[key] = property.column_type


    def execute(self, file_path: str, target_dir: str, sheet=DEFAULT_SHEET, global_param: dict = None,
                on_finish_one=default_on_finish_one):
        """
        :param on_finish_one: callback, on_finish_one(number: int, total_cnt: int, success: bool, exception: BaseException)
        """
        if global_param is None:
            global_param = {}
        self._check_and_set_default_global_param(global_param)
        data_list = self.__read(file_path, sheet)
        merged_data = self._merge_data(data_list=data_list, merge_fun_dict=self.merge_fun_dict)

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
                doc = self._process(data_list=sub_data_list, template_doc=template_doc, global_param=global_param,
                                    merged_data=merged_data)
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
        raw_datas = pandas.read_excel(file_path, sheet, dtype=self.dtype)
        data_list = []
        for i in range(raw_datas.shape[0]):
            row = raw_datas.iloc[i]
            raw_data = {}

            valid = True
            for k, filed_property in self.filed_mapping.items():
                filed_property: ExcelFiledProperty
                col_idx = excel_title_to_index(filed_property.column_title)
                if col_idx >= len(row):
                    continue
                val = self._castValue(row[col_idx], filed_property.value_cast)
                if val is None:
                    if self.divide_key == k or filed_property.required:
                        valid = False
                        break
                if filed_property.value_cast is not None:
                    if not filed_property.valid_check(val):
                        valid = False
                        break
                if val is not None:
                    raw_data[k] = val
            if valid:
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
    def _process(self, data_list: list, template_doc: Document, global_param: dict, merged_data: dict) -> Document:
        pass

    def _merge_data(self, data_list: List[dict], merge_fun_dict: dict) -> Dict[str, dict]:
        """
        :param merge_fun_dict mappingKey -> merge_fun
               merge_fun:merge_fun(mappingDataList) -> obj
        """
        merged_dict = {}
        if data_list is None:
            return merged_dict
        custom_merge_list_dict = {}
        for data in data_list:
            for k, v in data.items():
                if merge_fun_dict is not None and k in merge_fun_dict:
                    if k not in custom_merge_list_dict:
                        custom_merge_list_dict[k] = []
                    custom_merge_list_dict[k].append(v)
                else:
                    if v is None:
                        continue
                    else:
                        if k in merged_dict:
                            continue
                        else:
                            merged_dict[k] = v

        for k, mapping_list in custom_merge_list_dict.items():
            data = merge_fun_dict[k](mapping_list)
            if data is not None:
                merged_dict[k] = data
        return merged_dict

    def _get_file_name(self, key):
        return f'{key}.docx'

    def _castValue(self, value, cast):
        if cast is None:
            return value
        return cast(value)

    def _save(self, doc: Document, file_path: str):
        doc.save(file_path)


class MultiReportGenerator(metaclass=ABCMeta):
    """
       Abstract excel to excel or word report generator
       Input: one excel file
       Output: multi word files
       """

    # key -> cast
    #

    # args: number([0:N)), totalCount(N), success:bool, exception: BaseException

    def __init__(self, template_path: str, filed_mapping: Dict[str, ExcelFiledProperty], divide_key: str,
                 doc_global_data_param_config_list: List[DocGlobalParamConfig] = None,
                 merge_fun_dict: dict = None):
        """
        :param template_path: path to template docx file
        :param filed_mapping:
        :param divide_key: unique key

        dtype: Data type for data or columns. E.g. {‘a’: np.float64, ‘b’: np.int32} Use object to preserve data as stored in Excel and not interpret dtype. If converters are specified, they will be applied INSTEAD of dtype conversion.
            https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
        """
        super().__init__()
        self.template_path = template_path
        self.filed_mapping = filed_mapping
        self.divide_key = divide_key
        self.doc_global_data_param_config_list = doc_global_data_param_config_list
        self.merge_fun_dict = merge_fun_dict
        self.dtype = {}
        for key, property in filed_mapping.items():
            if property is not None and property.column_type is not None:
                self.dtype[key] = property.column_type

    def execute(self, file_path: str, target_dir: str, sheet=DEFAULT_SHEET, global_param: dict = None,
                on_finish_one=default_on_finish_one):
        """
        :param on_finish_one: callback, on_finish_one(number: int, total_cnt: int, success: bool, exception: BaseException)
        """
        if global_param is None:
            global_param = {}
        self._check_and_set_default_global_param(global_param)
        data_list = self.__read(file_path, sheet)
        merged_data = self._merge_data(data_list=data_list, merge_fun_dict=self.merge_fun_dict)

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
                doc = self._process(data_list=sub_data_list, template_doc=template_doc, global_param=global_param,
                                    merged_data=merged_data)
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
        raw_datas = pandas.read_excel(file_path, sheet, dtype=self.dtype)
        data_list = []
        for i in range(raw_datas.shape[0]):
            row = raw_datas.iloc[i]
            raw_data = {}

            valid = True
            for k, filed_property in self.filed_mapping.items():
                filed_property: ExcelFiledProperty
                col_idx = excel_title_to_index(filed_property.column_title)
                if col_idx >= len(row):
                    continue
                val = self._castValue(row[col_idx], filed_property.value_cast)
                if val is None:
                    if self.divide_key == k or filed_property.required:
                        valid = False
                        break
                if val is not None:
                    raw_data[k] = val
            if valid:
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
    def _process(self, data_list: list, template_doc: Document, global_param: dict, merged_data: dict) -> Document:
        pass

    def _merge_data(self, data_list: List[dict], merge_fun_dict: dict) -> Dict[str, dict]:
        """
        :param merge_fun_dict mappingKey -> merge_fun
               merge_fun:merge_fun(mappingDataList) -> obj
        """
        merged_dict = {}
        if data_list is None:
            return merged_dict
        custom_merge_list_dict = {}
        for data in data_list:
            for k, v in data.items():
                if merge_fun_dict is not None and k in merge_fun_dict:
                    if k not in custom_merge_list_dict:
                        custom_merge_list_dict[k] = []
                    custom_merge_list_dict[k].append(v)
                else:
                    if v is None:
                        continue
                    else:
                        if k in merged_dict:
                            continue
                        else:
                            merged_dict[k] = v

        for k, mapping_list in custom_merge_list_dict.items():
            data = merge_fun_dict[k](mapping_list)
            if data is not None:
                merged_dict[k] = data
        return merged_dict

    def _get_file_name(self, key):
        return f'{key}.docx'

    def _castValue(self, value, cast):
        if cast is None:
            return value
        return cast(value)

    def _save(self, doc: Document, file_path: str):
        doc.save(file_path)