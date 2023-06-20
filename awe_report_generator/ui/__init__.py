from PyQt5.QtWidgets import QWidget


class BaseUIConfig:
    def __init__(self, name: str, qwidget: QWidget):
        self.name = name
        self.qwidget = qwidget
