from typing import List, Dict
import threading

from PyQt5.QtWidgets import (QWidget,
                             QVBoxLayout,
                             QPushButton, QGridLayout, QLineEdit, QLabel, QFileDialog, QMessageBox, QProgressBar,
                             )


class ProcessorUIParam:
    def __init__(self, title, processor: ReportGenerator) -> None:
        self.title = title
        self.processor = processor


class ProcessorQWidget(QWidget):
    def __init__(self, processor_ui_param: ProcessorUIParam):
        super().__init__()
        self.processor_ui_param = processor_ui_param
        self.input_file_path = None
        self.sheet = None
        self.output_dir_path = None
        self.global_param_widget_dict = {}
        self.msg_box = QMessageBox()
        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        grid.setSpacing(1)

        self.setLayout(grid)
        self.setGeometry(300, 300, 600, 100)
        self.setWindowTitle(self.processor_ui_param.title)

        doc_global_data_param_config_list = self.processor_ui_param.processor.doc_global_data_param_config_list
        i = 0

        if doc_global_data_param_config_list is not None:
            for param_config in doc_global_data_param_config_list:
                qlabel = QLabel()
                if param_config.required:
                    qlabel.setText(f'{param_config.title}*')
                else:
                    qlabel.setText(f'{param_config.title}')
                qline_edit = QLineEdit()
                self.global_param_widget_dict[param_config.filed] = qline_edit
                qline_edit.setText(param_config.default_value)
                grid.addWidget(qlabel, i, 0)
                grid.addWidget(qline_edit, i, 1)
                i = i + 1

        sheet_qlabel = QLabel('请输入表格sheet名称')
        grid.addWidget(sheet_qlabel, i, 0)

        self.sheet_edit = QLineEdit('空分质量部抽检')
        grid.addWidget(self.sheet_edit, i, 1)
        i = i + 1

        input_file_btn = QPushButton()
        input_file_btn.setText('选择输入文件(xlsx)*')
        input_file_btn.clicked.connect(self.select_input_file)
        self.input_file_qlabel = QLabel()

        grid.addWidget(input_file_btn, i, 0)
        grid.addWidget(self.input_file_qlabel, i, 1)
        i = i + 1

        output_dir_btn = QPushButton()
        output_dir_btn.setText('选择输出文件夹*')
        output_dir_btn.clicked.connect(self.select_output_dir)
        grid.addWidget(output_dir_btn, i, 0)
        self.output_dir_qlabel = QLabel()
        grid.addWidget(self.output_dir_qlabel, i, 1)
        i = i + 1

        submit_btn = QPushButton()
        submit_btn.setText('提交')
        submit_btn.clicked.connect(self.on_submit)
        grid.addWidget(submit_btn, i, 1)
        i = i + 1

        self.progress_bar = QProgressBar()
        grid.addWidget(self.progress_bar, i, 0, 1, 3)
        i = i + 1

    def select_input_file(self):
        self.input_file_path, _ = QFileDialog.getOpenFileName(self, "选择输入文件",
                                                              None,
                                                              "*.xlsx")
        self.input_file_qlabel.setText(self.input_file_path)

    def select_output_dir(self):
        self.output_dir_path = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "./")
        self.output_dir_qlabel.setText(self.output_dir_path)

    def on_submit(self):
        global_param = {}
        doc_global_data_param_config_list = self.processor_ui_param.processor.doc_global_data_param_config_list

        if doc_global_data_param_config_list is not None:
            for param_config in doc_global_data_param_config_list:
                qline_edit: QLineEdit = self.global_param_widget_dict[param_config.filed]
                val = qline_edit.text()
                if param_config.required and (val is None or val == ''):
                    self.alert(f'请输入必填参数:{param_config.title}')
                    return
                if val is not None:
                    global_param[param_config.filed] = val

        if self.input_file_path is None or self.input_file_path == '':
            self.alert(f'请选择输入文件')
            return

        if self.output_dir_path is None or self.output_dir_path is None:
            self.alert(f'请选择输出文件夹')
            return

        sheet = self.sheet_edit.text()
        if sheet is None or sheet == '':
            self.alert('请输入Sheet名称')
            return

        self.work_thread = threading.Thread(name=f'Processor-{processor_ui_param.title}', target=self.processor_ui_param.processor.execute, args=(
            self.input_file_path, self.output_dir_path, sheet, global_param, self.on_finish_one))
        # self.work_thread = threading.Thread(name='Processor', target=self.process)
        self.work_thread.start()

    def alert(self, msg):
        self.msg_box.setText(msg)
        self.msg_box.show()

    def on_finish_one(self, number: int, total_cnt: int, success: bool, exception: BaseException):
        print(f'{number}/{total_cnt}, success={success}, e={exception}')
        if number == 1:
            self.progress_bar.reset()
            self.progress_bar.setRange(1, total_cnt)

        self.progress_bar.setValue(number)
