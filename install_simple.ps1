conda activate p38
pyinstaller -F --noconsole --hidden-import openpyxl.cell._writer app.py
