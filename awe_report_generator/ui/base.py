from typing import List

from PyQt5.QtWidgets import (QWidget,
                             QGridLayout,
                             QPushButton,
                             )

from . import ProcessorUIConfig
from .processor_ui import ProcessorQWidget


class HomePageQWidget(QWidget):
    def __init__(self, processor_ui_config_list: List[ProcessorUIConfig]):
        super().__init__()
        self.processor_ui_config_list = processor_ui_config_list
        self.processor_widget_list = []

        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        grid.setSpacing(0)

        for processor_ui_config in self.processor_ui_config_list:
            btn = self.build_button(processor_ui_config)
            processor_widget = ProcessorQWidget(processor_ui_config=processor_ui_config)

            self.processor_widget_list.append(processor_widget)
            btn.clicked.connect(processor_widget.show)
            grid.addWidget(btn)

        self.setLayout(grid)
        self.setGeometry(300, 300, 600, 100)
        self.setWindowTitle('文档自动生成器')

    def build_button(self, processor_ui_config: ProcessorUIConfig):
        btn = QPushButton(processor_ui_config.name, self)
        btn.setToolTip(f'点击进入 <b>{processor_ui_config.name}</b>')
        btn.resize(btn.sizeHint())
        return btn
