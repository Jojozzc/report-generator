from typing import List

from PyQt5.QtWidgets import (QWidget,
                             QGridLayout,
                             QPushButton,
                             )

from . import BaseUIConfig


class HomePageQWidget(QWidget):
    def __init__(self, ui_config_list: List[BaseUIConfig]):
        super().__init__()
        self.ui_config_list = ui_config_list

        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        grid.setSpacing(0)

        for ui_config in self.ui_config_list:
            btn = self.build_button(ui_config)

            btn.clicked.connect(ui_config.qwidget.show)
            grid.addWidget(btn)

        self.setLayout(grid)
        self.setGeometry(300, 300, 600, 100)
        self.setWindowTitle('文档自动生成器')

    def build_button(self, ui_config: BaseUIConfig):
        btn = QPushButton(ui_config.name, self)
        btn.setToolTip(f'点击进入 <b>{ui_config.name}</b>')
        btn.resize(btn.sizeHint())
        return btn
