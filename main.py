# coding = utf-8
# main.py

import sys
from PyQt5.QtWidgets import QApplication
from qt import *

# Initilize graphic user interface
app = QApplication(sys.argv)
app_win = AppWindow()
sys.exit(app.exec_())
