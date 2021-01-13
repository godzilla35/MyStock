import sys
from kiwoom.kiwoom import *
from PyQt5.QtWidgets import *

class UI_class():
    def __init__(self):
        print("UI class")
        self.app = QApplication(sys.argv)
        self.k = Kiwoom()

        self.app.exec_() #이벤트 루프 실행