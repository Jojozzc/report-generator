from awe_report_generator.core.base import ReportGenerator



class BaseUIConfig:
    def __init__(self, name: str, qwidget: QWidget):
        self.name = name
        self.qwidget = qwidget
