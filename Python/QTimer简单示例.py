import sys

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer, QObject

class AA(QObject):
    def __init__(self):
        super(QObject, self).__init__()
        self.time = QTimer(self)
        self.time.start(2000)
        self.time.timeout.connect(self.abc)

    def abc(self):
        print('33333333')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    d = AA()
    sys.exit(app.exec_())