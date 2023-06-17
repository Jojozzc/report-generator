import sys

from PyQt5.QtWidgets import (QApplication,
                             )
from ui import ProcessorQWidget


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ProcessorQWidget()
    sys.exit(app.exec_())