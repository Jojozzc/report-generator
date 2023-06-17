#!/usr/bin/python3
# -*- coding: utf-8 -*-
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QWidget,
                             QVBoxLayout,
                             QPushButton,
                             )


class ProcessorQWidget(QWidget):


    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        buttons = self.businessButtons()

        vbox = QVBoxLayout()
        vbox.setSpacing(0)

        for btn in buttons:
            vbox.addWidget(btn)

        self.setLayout(vbox)
        self.setGeometry(300, 300, 600, 300)
        self.setWindowTitle('表面结果生成器')
        self.show()

    def businessButtons(self) -> list:
        buttons = []
        btn1 = QPushButton('委托生成器', self)
        btn1.setToolTip('点击进入 <b>委托生成器</b>')
        btn1.resize(btn1.sizeHint())

        btn2 = QPushButton('射线结果生成器', self)
        btn2.setToolTip('点击进入 <b>射线结果生成器</b>')
        btn2.resize(btn2.sizeHint())

        btn3 = QPushButton('表面结果生成器', self)
        btn3.setToolTip('点击进入 <b>表面结果生成器</b>')
        btn3.resize(btn3.sizeHint())

        buttons.append(btn1)
        buttons.append(btn2)
        buttons.append(btn3)

        return buttons
