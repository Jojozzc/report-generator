from PyQt5.QtCore import QSettings

ORG_NAME = 'JSoft'
APP_NAME = 'AwesomeReportGenerator'

APP_Q_SETTINGS = QSettings(ORG_NAME, APP_NAME)


def put_setting(biz_code: str, key: str, value: str):
    APP_Q_SETTINGS.setValue(_get_full_key(biz_code, key), value)
    APP_Q_SETTINGS.sync()


def get_setting(biz_code: str, key: str, default_value=None):
    return APP_Q_SETTINGS.value(_get_full_key(biz_code, key), default_value)


def _get_full_key(biz_code: str, key: str):
    return f'{biz_code}/{key}'
