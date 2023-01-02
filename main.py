import write_processor
import tkinter as tk
from tkinter import filedialog


def on_submit():
    submit_button.config(text='请稍后')
    submit_button.config(state=tk.DISABLED)
    input_file_path = './testData/气化数据库2.xlsx'
    template_path = './template/TEMPLATE.docx'

    write_processor.process(template_path, input_file_select,
                            title1_input.get(),
                            title2_input.get(), company_input.get(), customer_input.get(), method_input.get(),
                            standard_input.get(), output_dir)
    submit_button.config(text='提交')
    submit_button.config(state=tk.NORMAL)


def on_open_docx_file():
    global input_file_select
    input_file_select = filedialog.askopenfilename(title='选择输入文件', filetypes=[('excel表格', '*.xlsx')])
    file_select_display.config(text=f'已选择:{input_file_select}')
    print(f'input_file_select={input_file_select}')


def on_select_out_put_dir():
    global output_dir
    output_dir = filedialog.askdirectory(title='选择输出文件夹')
    output_display.config(text=f'已选择:{output_dir}')
    print(f'output_dir={output_dir}')


if __name__ == '__main__':
    root_window = tk.Tk()
    root_window.title('文档自动生成器')
    root_window.geometry('1000x500')

    title1_label = tk.Label(text='请输入工程名称:')
    title1_input = tk.Entry(root_window, textvariable=tk.StringVar(value='宁夏宝丰能源集团股份有限公司50万吨/年煤制烯烃项目配套甲醇工程'))

    title2_label = tk.Label(text='请输入单位工程名称:')
    title2_input = tk.Entry(root_window, textvariable=tk.StringVar(value='气化1'))

    company_label = tk.Label(text='请输入检测单位:')
    company_input = tk.Entry(root_window, textvariable=tk.StringVar(value='南京英派克检测有限责任公司'))

    customer_label = tk.Label(text='请输入委托单位:')
    customer_input = tk.Entry(root_window, textvariable=tk.StringVar(value='中化六建'))

    method_label = tk.Label(text='请输入检测方法:')
    method_input = tk.Entry(root_window, textvariable=tk.StringVar(value='RT'))

    standard_label = tk.Label(text='请输入检测标准:')
    standard_input = tk.Entry(root_window, textvariable=tk.StringVar(value='NB/T47013.2-2015'))

    file_select_button = tk.Button(root_window, text='请选择输入文件', command=on_open_docx_file)
    file_select_display = tk.Label(root_window, text='未选择')

    output_button = tk.Button(root_window, text='请选择输出文件夹', command=on_select_out_put_dir)
    output_display = tk.Label(root_window, text='未选择')

    title1_label.pack()
    title1_input.pack()
    title2_label.pack()
    title2_input.pack()
    company_label.pack()
    company_input.pack()
    customer_label.pack()
    customer_input.pack()
    method_label.pack()
    method_input.pack()
    standard_label.pack()
    standard_input.pack()

    file_select_button.pack()
    file_select_display.pack()
    output_button.pack()
    output_display.pack()
    submit_button = tk.Button(root_window, text="提交",
                              command=lambda: on_submit())
    submit_button.pack(side="bottom")

    root_window.mainloop()
