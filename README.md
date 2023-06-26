# 打包方法
目前使用New App，mode1和model2不再维护

## mode1
1. 代码入口：main.py
2. 模板：template/TEMPLATE.docx
3. 执行 pyinstaller -F main.py

## mode2
1. mode2_main.py
2. 模板：template/TEMPLATE_MODE2.docx
3. 执行 pyinstaller -F mode2_main.py

## New App
1. app.py
2. 模板：template/*
3. 执行 pyinstaller -F app.py
