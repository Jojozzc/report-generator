from docx.enum.table import WD_TABLE_ALIGNMENT


class DocCellStyle:
    def __init__(self, font_cn = None, alignment=WD_TABLE_ALIGNMENT.CENTER, font_size=None):
        self.font_cn = font_cn
        self.alignment = alignment
        self.font_size = font_size