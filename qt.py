# coding = utf-8

import lighter
import time
from PyQt5 import QtCore, sip
from PyQt5.QtWidgets import (QWidget, QGridLayout, QLabel, QLineEdit, QPushButton, QApplication)
from functions import *
 
class AppWindow(QWidget):

    # Options
    # structure-file & price-file options
    output_file_name = 'bomdaji' + str(time.strftime('%Y%m%d', time.gmtime()))
    options = {
        (0, 0): '组合构成文件名', (0, 1): '组合构成',          # 组合结构文件名
        (1, 0): '组合构成表名',   (1, 1): 'Sheet1',           # 组合结构文件表名
        (2, 0): '父物料编码列号', (2, 1): 'F',                # 列号 - 父项物料编码
        (3, 0): '项次列号',       (3, 1): 'K',                # 列号 - 项次
        (4, 0): '子物料编码列号', (4, 1): 'L',                # 列号 - 子项物料编码
        (5, 0): '用量分子列号',   (5, 1): 'Q',                # 列号 - 用量分子
        (6, 0): '用量分母列号',   (6, 1): 'R',                # 列号 - 用量分母
        (7, 0): '物料名称列号',   (7, 1): 'G',                # 列号 - 物料名称
        (8, 0): '规格型号列号',   (8, 1): 'H',                # 列号 - 规格型号
        (0, 2): '单价文件名',     (0, 3): '单价',             # 价格文件名
        (1, 2): '组合构成表名',   (1, 3): 'Sheet1',           # 价格文件表名
        (2, 2): '物料编码列号',   (2, 3): 'D',                # 列号 - 物料编码
        (3, 2): '价格列号',       (3, 3): 'I',                # 列号 - 价格
        (4, 2): '含税价格列号',   (4, 3): 'J',                # 列号 - 含税价格
        (9, 0): '',              (10, 0): '',                # 占位符
        (11, 0): '导出文件名',    (11, 1): output_file_name   # 导出文件名
    }
    qle = {}

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        # initialize file options & line editor
        for prow, pcol in self.options:
            if pcol % 2 == 0:
                qlb = QLabel(self.options[prow, pcol])
                grid.addWidget(qlb, prow, pcol)
            else:
                self.qle[prow, pcol] = QLineEdit()
                self.qle[prow, pcol].setText(self.options[prow, pcol])
                self.qle[prow, pcol].setPlaceholderText(self.options[prow, pcol])
                # self.qle[prow, pcol].textChanged[str].connect(self.onOptChanged)
                grid.addWidget(self.qle[prow, pcol], prow, pcol)
        # initialize button
        launch_btn = QPushButton("生成价格表")
        # connect to readyLaunch method
        launch_btn.clicked.connect(self.readyLaunch)
        grid.addWidget(launch_btn, 11, 3)
        # readmeL1 = QLabel('1. 将组合结构文件及单价文件与本程序放到相同目录下。', self)
        # readmeL2 = QLabel('2. 确认各列号与对应的数据保持一致。', self)
        # readmeL3 = QLabel('3. 点击按钮"生成价格表"导出bom表价格清单。', self)
        # readmeL1.move(250, 150)
        # readmeL2.move(250, 170)
        # readmeL3.move(250, 190)

        self.move(300, 150)
        self.setWindowTitle('BOM表价格整合工具 - HJ No.7')
        self.show()

    def onOptChanged(self, text):
        self.options[self.tempRow, self.tempCol] = text

    def readyLaunch(self):
        # update edit-bar values
        for r, c in self.qle:
            self.options[r, c] = self.qle[r, c].text()
        # connect to launch method in lighter.py
        lighter.launch(self.options)