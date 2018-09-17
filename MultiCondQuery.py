# !python3.6
# -*- coding: utf-8 -*-

import sys
import PyQt5
import pandas as pd
import pyperclip

from PyQt5 import QtCore
from PyQt5.QtWidgets import QShortcut, QSizePolicy, QTableView
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QApplication, QPushButton
from PyQt5.QtWidgets import QFormLayout, QFileDialog, QComboBox, QLineEdit
from qtpandas.models.DataFrameModel import DataFrameModel
from qtpandas.views.DataTableView import DataTableWidget

class ResultWindow(QWidget):
	def __init__(self):
		super(self.__class__, self).__init__()
		self.setup_ui()
		
	def setup_ui(self):
		self.setWindowTitle("查詢結果")
		self.setFont(PyQt5.QtGui.QFont("Times New Roman", 11))
		
		self.buttonCopy = QPushButton("複製查詢結果 (Ctrl+C)")
		self.buttonCopy.clicked.connect(self.copy)
		self.labelInfo = QLabel("")
		self.shortcut = QShortcut(PyQt5.QtGui.QKeySequence("Ctrl+C"), self)
		self.shortcut.activated.connect(self.copy)

		self.formLayout = QFormLayout()
		self.formLayout.addRow(self.buttonCopy, self.labelInfo)

		self.widget = DataTableWidget()
		self.widget.setButtonsVisible(False)
		self.widget.setMinimumSize(600, 600)

		self.vbox = QVBoxLayout()
		self.vbox.addLayout(self.formLayout)
		self.vbox.addWidget(self.widget)
		self.setLayout(self.vbox)

	def update(self, data):
		self.df = data
		self.labelInfo.setText("")
		self.model = DataFrameModel()

		self.widget.resize(self.widget.sizeHint())
		self.widget.show()
		self.widget.setViewModel(self.model)
		self.model.setDataFrame(data)

	def copy(self):
		count = 0
		indexes = self.widget.tableView.selectionModel().selectedRows()
		if len(indexes) > 0:
			rows = [index.row() for index in indexes]
			count = len(indexes)
		else:
			rows = [i for i in range(self.widget.tableView.model().rowCount())]
			count = self.widget.tableView.model().rowCount()
		s = ''
		for i in rows:
			r = self.df.iloc[[i]]
			for v in r.values[0].astype(str):
				s += (v + '\t')
			s = s[:-1] + '\n'
		pyperclip.copy(s[:-1])
		self.labelInfo.setText(str(count) + " 筆資料已複製至剪貼簿")

class MainWindow(QWidget):
	def __init__(self):
		super(self.__class__, self).__init__()
		self.conditions = []
		self.contents = []
		self.MAX_COND = 4
		self.resultWindow = ResultWindow()
		self.setup_ui()
		self.show()
		self.load_file()

	def setup_ui(self):
		self.setWindowTitle("Excel多重條件篩選工具")
		self.setFixedWidth(600)
		self.setFont(PyQt5.QtGui.QFont("Times New Roman", 11))

		self.buttonLoad = QPushButton("載入")
		self.buttonLoad.clicked.connect(self.load_file)
		self.labelInfo = QLabel("")

		self.buttonReset = QPushButton("清除 (Esc)")
		self.buttonReset.clicked.connect(self.reset)

		self.buttonQuery = QPushButton("查詢 (Enter)")
		self.buttonQuery.setEnabled(False)
		self.buttonQuery.clicked.connect(self.query)

		funcformLayout = QFormLayout()
		funcformLayout.addRow(self.buttonLoad, self.labelInfo)
		funcformLayout.addRow(self.buttonReset, self.buttonQuery)
		
		self.vbox = QVBoxLayout()
		self.vbox.addLayout(funcformLayout)

		self.setLayout(self.vbox)

	def load_file(self):
		options = QFileDialog.Options()
		filename, _ = QFileDialog.getOpenFileName(self, "載入檔案", "", "Excel Files (*.xlsx)", options=options)
		if filename:
			self.df = pd.read_excel(filename, encoding='utf-8', dtype=str)
			self.labelInfo.setText(filename.split('/')[-1])
			self.buttonQuery.setEnabled(True)
			condformLayout = QFormLayout()
			
			for i in range(self.MAX_COND):
				combobox = QComboBox()
				combobox.addItems(list(self.df))
				line = QLineEdit()
				self.conditions.append(combobox)
				self.contents.append(line)
				condformLayout.addRow(self.conditions[i], self.contents[i])
			self.vbox.addLayout(condformLayout)
			self.setLayout(self.vbox)

	def _get_logic(self, cond, keywords):
		keys = []
		if "and" in keywords:
			logic = True
			keys = [s.strip() for s in keywords.split("and")]
			for k in keys:
				logic = logic & self.df[str(cond)].str.contains(k)
		elif "or" in keywords:
			logic = False
			keys = [s.strip() for s in keywords.split("or")]
			for k in keys:
				logic = logic | self.df[str(cond)].str.contains(k)
		else:
			logic = self.df[str(cond)].str.contains(keywords)
		return logic

	def query(self):
		# AND logic between conditions, AND or OR logic in one condition
		logic = True
		for i in range(0, len(self.conditions)):
			cond = self.conditions[i].currentText()
			keyword = self.contents[i].text()
			logic = logic & self._get_logic(cond, keyword)
		result = self.df[logic]
		self.resultWindow.move(50,50)
		self.resultWindow.update(result)
		self.resultWindow.show()

	def reset(self):
		for i in range(len(self.contents)):
			self.contents[i].setText('')

	def keyPressEvent(self, e):
		if e.key() == QtCore.Qt.Key_Escape:
			self.reset()
		elif e.key() in [QtCore.Qt.Key_Enter, QtCore.Qt.Key_Return]:
			self.query()

	def closeEvent(self, event):
		self.resultWindow.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    sys.exit(app.exec_())