# 打包方法
1. 程序入口：app.py
2. 模板文件夹：template/*
3. 执行 pyinstaller -F --noconsole --hidden-import openpyxl.cell._writer app.py

# 项目介绍
主要功能：将批量的Excel文件汇总成多个Word报表

## 核心依赖库
- PyQt5：用于UI开发
- xlsxwriter：用于Excel文件的读写
- docx：用于Word文件的读写

## 主要文件
- 根目录
app.py文件为入口，可以直接运行该文件进行debug。
阅读代码时建议从app.py开始，并逐步深入到各个模块。

- awe_report_generotor：该文件夹包含了程序除了app.py外的全部代码。
- awe_report_generotor/ui：主要存放UI界面相关的代码。
- awe_report_generotor/app：app层面的代码，目前主要是关于PYQT框架的存储配置。
- awe_report_generotor/core：包含了业务无关的核心代码，抽象最核心的部分。
- awe_report_generotor/biz：该文件夹主要包含业务自定义的数据格式转换类。
- awe_report_generotor/utils：该文件夹包含了工具类代码。
- template文件夹：存放Word报表的模板文件。属于重要的资源文件。
