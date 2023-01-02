import write_processor

if __name__ == '__main__':
    input_file_path = './testData/气化数据库2.xlsx'
    template_path = './template/TEMPLATE.docx'
    write_processor.process(template_path, input_file_path,
                            '宁夏宝丰能源集团股份有限公司50万吨/年煤制烯烃项目配套甲醇工程',
                            '气化', '南京英派克检测有限责任公司', '中化六建', 'RT', 'NB/T47013.2-2015', './testData')
