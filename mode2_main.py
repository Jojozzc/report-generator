import sys

from mode2 import write_processor_mode2
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import threading
import datetime


class AppStdout:
    backup_stdout = None
    backup_stderr = None
    text_ui: tk.Text = None

    def __init__(self, text_ui):
        self.text_ui = text_ui
        self.backup_stdout = sys.stdout
        self.backup_stderr = sys.stderr
        sys.stdout = self
        sys.stderr = self

    def write(self, info):
        self.text_ui.insert('end', info)
        self.text_ui.update()
        self.text_ui.see(tk.END)

    def close(self):
        sys.stdout = self.backup_stdout
        sys.stderr = self.backup_stderr


def process_callback(total_cnt, fin_cnt, fail_cnt, problem_cnt):
    progress_hint_label.config(text=f'({fin_cnt}/{total_cnt}) 失败文件数:{fail_cnt},需要关注的文件数:{problem_cnt}')
    if total_cnt <= 0:
        progressbar['value'] = 100
    elif total_cnt <= fin_cnt:
        progressbar['value'] = 100
    else:
        cnt_val = int(fin_cnt / total_cnt * 100)
        cnt_val = min(cnt_val, 99)
        progressbar['value'] = cnt_val
    root_window.update()


def async_process(template_path: str, input_file_path: str, title1: str, title2: str, customer: str,
                  method: str, target_dir: str):
    try:
        write_processor_mode2.process(template_path, input_file_path, title1, title2, customer, method,
                            target_dir, process_callback)
        hint_label.config(text='执行成功！')
    except BaseException as e:
        hint_label.config(text='执行失败！错误:' + str(e))
    finally:
        submit_button.config(text='提交')
        submit_button.config(state=tk.NORMAL)
    print(f'---执行完成 {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}---')
    root_window.update()


def on_submit():

    std_output_text.delete(1.0, tk.END)
    print(f'---开始执行 {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}---')
    submit_button.config(text='请稍后')
    hint_label.config(text='执行中，请稍后')
    progress_hint_label.config(text='')
    submit_button.config(state=tk.DISABLED)
    template_path = './template/TEMPLATE_MODE2.docx'
    progressbar['maximum'] = 100
    progressbar['value'] = 0
    root_window.update()
    try:
        work_thread = threading.Thread(name='Processor', target=async_process, args=(template_path, input_file_select,
                                                                                     title1_input.get(),
                                                                                     title2_input.get(),
                                                                                     customer_input.get(),
                                                                                     method_input.get(),
                                                                                     output_dir))
        work_thread.start()
        hint_label.config(text='提交成功！')
    except Exception as e:
        print(f'执行错误！{e}')
        print(e)
        hint_label.config(text='执行出错')
        submit_button.config(text='提交')
        submit_button.config(state=tk.NORMAL)
        progressbar['value'] = 0


def on_open_docx_file():
    global input_file_select
    input_file_select = filedialog.askopenfilename(title='选择输入文件', filetypes=[('excel表格', '*.xlsx')])
    file_select_display.config(text=f'已选择:{input_file_select}')


def on_select_out_put_dir():
    global output_dir
    output_dir = filedialog.askdirectory(title='选择输出文件夹')
    output_display.config(text=f'已选择:{output_dir}')


if __name__ == '__main__':
    root_window = tk.Tk()
    root_window.title('文档自动生成器')
    root_window.geometry('1000x700')

    title1_label = tk.Label(root_window, text='请输入工程名称:')
    title1_input = tk.Entry(root_window, textvariable=tk.StringVar(value='川西气田雷口坡组气藏开发建设项目脱硫站工程5#脱硫站'))

    title2_label = tk.Label(root_window, text='请输入单位工程名称:')
    title2_input = tk.Entry(root_window, textvariable=tk.StringVar(value='测试'))

    customer_label = tk.Label(root_window, text='请输入委托单位:')
    customer_input = tk.Entry(root_window, textvariable=tk.StringVar(value='中国石化第四建设有限公司'))

    method_label = tk.Label(root_window, text='请输入检测方法:')
    method_input = tk.Entry(root_window, textvariable=tk.StringVar(value='RT'))


    file_select_button = tk.Button(root_window, text='请选择输入文件', command=on_open_docx_file)
    file_select_display = tk.Label(root_window, text='未选择')

    output_button = tk.Button(root_window, text='请选择输出文件夹', command=on_select_out_put_dir)
    output_display = tk.Label(root_window, text='未选择')

    progressbar = ttk.Progressbar(root_window)
    progress_hint_label = tk.Label(root_window, text='')
    hint_label = tk.Label(root_window, text='')
    std_output_text = tk.Text(root_window, height=10, width=100)

    title1_label.pack()
    title1_input.pack()
    title2_label.pack()
    title2_input.pack()
    customer_label.pack()
    customer_input.pack()
    method_label.pack()
    method_input.pack()

    file_select_button.pack()
    file_select_display.pack()
    output_button.pack()
    output_display.pack()
    progressbar.pack()
    progress_hint_label.pack()
    hint_label.pack()

    submit_button = tk.Button(root_window, text="提交",
                              command=on_submit)
    # submit_button.pack(side="bottom")
    submit_button.pack()
    std_output_text.pack()
    app_std_out = AppStdout(std_output_text)

    root_window.mainloop()

    app_std_out.close()
