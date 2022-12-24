import docx
import pandas


if __name__ == '__main__':
    # file_path = './testData/EPK-BFJC-MM-3115-AZ(02)-RT-039.docx'
    input_file_path = './testData/气化数据库.xlsx'
    template_path = './template/TEMPLATE.docx'


    raw_data = pandas.read_excel(input_file_path, 'Sheet1')

    data_map = {}

    for i in range(raw_data.shape[0]):
        index = raw_data.iloc[i, 2]
        if index not in data_map:
            data_map[index] = []
        data_map[index].append(raw_data.iloc[i])
    xx = raw_data.iloc[1,2]



    for key in data_map.keys():
        values = data_map[key]




    doc = docx.Document(template_path)
    doc.paragraphs[1].text = doc.paragraphs[1].text + '宁夏宝丰能源集团股份有限公司50万吨/年煤制烯烃项目配套甲醇工程'
    doc.paragraphs[2].text = doc.paragraphs[2].text + '气化1'
    table = doc.tables[0]
    table.cell(0, 2).text = '南京英派克检测有限责任公司'
    table.cell(0, 11).text = 'BFJC-QH-3117-AZ(02)-RT-01408'

    table.cell(1, 2).text = '中化六建'
    table.cell(1, 11).text = '2022.12.02'


    table.cell(2, 2).text = 'RT'
    table.cell(2, 8).text = 'NB/T47013.2-2015'
    table.cell(2, 14).text = 'III'





    # # 检测单位
    # table.cell(0, 3).text = '南京英派克检测有限责任公司'
    # # table.cell(0, 3).text = 'BFJC-QH-3117-AZ(02)-RT-01408'
    # table.cell(1, 1).text = '中化六建'
    # table.cell(1, 3).text = '2022.12.02'
    # table.cell(2, 1).text = 'RT'
    # table.cell(2, 3).text = 'NB/T47013.2-2015'
    # table.cell(2, 5).text = 'III'

    doc.save('./testData/testDoc.docx')

    pass