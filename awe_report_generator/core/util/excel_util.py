from typing import Dict

import pandas

from awe_report_generator.core.base import ExcelFiledProperty
from awe_report_generator.core.util import excel_title_to_index


def read_data_list(file_path: str, sheet, filed_mapping: Dict[str, ExcelFiledProperty]):
    dtype = {}
    for key, property in filed_mapping.items():
        if property is not None and property.column_type is not None:
            dtype[key] = property.column_type
    raw_datas = pandas.read_excel(file_path, sheet, dtype=dtype)
    data_list = []
    for i in range(raw_datas.shape[0]):
        row = raw_datas.iloc[i]
        raw_data = {}

        valid = True
        for k, filed_property in filed_mapping.items():
            filed_property: ExcelFiledProperty
            col_idx = excel_title_to_index(filed_property.column_title)
            if col_idx >= len(row):
                continue
            val = _castValue(row[col_idx], filed_property.value_cast)
            if val is None:
                if filed_property.required:
                    valid = False
                    break
            if filed_property.valid_check is not None:
                if not filed_property.valid_check(val):
                    valid = False
                    break
            if val is not None:
                raw_data[k] = val
        if valid:
            data_list.append(raw_data)

    return data_list


def _castValue(value, cast):
    if cast is None:
        return value
    return cast(value)
