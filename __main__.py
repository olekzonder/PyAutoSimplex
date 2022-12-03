import sys
import simplex
import docxgen
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QWidget,
    QLabel,
    QTableWidget,
    QHeaderView,
    QComboBox,
    QTableWidgetItem,
    QFileDialog,
    QMessageBox
)
from PyQt5.QtCore import *

class Window(QWidget):

    
    def __init__(self,parent=None):
        super().__init__(parent)
        self.targetCount = 2
        self.constraintCount=1
        self.setWindowTitle("Симплекс-метод")
        outerLayout = QVBoxLayout()
        outerLayout.addWidget(QLabel(text='Целевая функция:',alignment=Qt.AlignCenter))

        self.targetTable = QTableWidget()
        self.targetTable.setRowCount(1)
        self.targetTable.setColumnCount(self.targetCount)
        self.targetTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)])
        self.targetTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.targetTable.verticalHeader().hide()
        self.targetTable.setMaximumHeight(75)
        for i in range(self.targetCount):
            self.targetTable.setItem(0,i,QTableWidgetItem(str(0)))
        outerLayout.addWidget(self.targetTable)

        self.addTButton = QPushButton(text='+')
        self.delTButton = QPushButton(text='-')
        TButtonLayout = QHBoxLayout()
        TButtonLayout.addWidget(self.addTButton)
        self.addTButton.clicked.connect(self.addT)
        self.delTButton.clicked.connect(self.delT)
        TButtonLayout.addWidget(self.delTButton)
        outerLayout.addLayout(TButtonLayout)

        self.comboBox = QComboBox()
        self.comboBox.addItems(["min","max"])
        self.type = self.comboBox.currentText()
        self.comboBox.currentTextChanged.connect(self.setType)
        outerLayout.addWidget(self.comboBox)

        outerLayout.addWidget(QLabel(text='Ограничения:',alignment=Qt.AlignCenter))
        self.constraintTable = QTableWidget()
        self.constraintTable.setRowCount(self.constraintCount)
        self.constraintTable.setColumnCount(self.targetCount+2)
        self.constraintTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)]+['']+["z"])
        for i in range(self.targetCount+2):
            self.constraintTable.setItem(0,i,QTableWidgetItem(str(0)))
        lmBox = QComboBox()
        lmBox.addItems(['≥',"≤"])
        self.constraintTable.setCellWidget(0,self.targetCount,lmBox)
        self.constraintTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        outerLayout.addWidget(self.constraintTable)
        
        self.addCButton = QPushButton(text='+')
        self.delCButton = QPushButton(text='-')
        CButtonLayout = QHBoxLayout()
        CButtonLayout.addWidget(self.addCButton)
        self.addCButton.clicked.connect(self.addC)
        self.delCButton.clicked.connect(self.delC)
        CButtonLayout.addWidget(self.delCButton)
        outerLayout.addLayout(CButtonLayout)

        self.startButton = QPushButton(text="ОК")
        self.startButton.clicked.connect(self.start)
        outerLayout.addWidget(self.startButton)
        self.setLayout(outerLayout)
        self.constraintTable.resizeRowsToContents()
        self.targetTable.resizeRowsToContents()
    
    def setType(self):
        self.type = self.comboBox.currentText()

    def start(self):
        self.function = []
        self.constraints = []
        self.resources = []
        self.lessThanList=[]
        for i in range(self.targetCount):
            self.function.append(float(self.targetTable.item(0,i).text()))
        for i in range(self.constraintCount):
            temp = []
            self.resources.append(float(self.constraintTable.item(i,self.targetCount+1).text()))
            lessThan =  self.constraintTable.cellWidget(i,self.targetCount).currentText()
            if lessThan == '≤':
                self.lessThanList.append(i)
            for j in range(self.targetCount):
                temp.append(float(self.constraintTable.item(i,j).text()))
            self.constraints.append(temp)
        filename = QFileDialog.getSaveFileName(self,'Сохраните файл',"Отчет.docx","Документ (*.docx)")
        filename = filename[0]
        if filename == '':
            return
        if ''.join(list(filename)[-5:]) != '.docx':
            filename= filename+'.docx'
        data = {}
        data['function'] = self.function
        data['constraints'] = self.constraints
        data['resources'] = self.resources
        data['type'] = self.type
        try:
            document = docxgen.DOCXGen(filename, data ,self.type,self.lessThanList)
            optimizer = simplex.Simplex(self.function,self.constraints,self.resources,type=self.type,lessThanList=self.lessThanList)
        except:
            QMessageBox.warning(self,"Ошибка","Ошибка при вводе значений")
            return
        try:
            solution = optimizer.optimize()
        except IndexError:
            QMessageBox.warning(self,"Ошибка","Ошибка при поиска решения - возможно, были введены неправильные значения")
            return
        except RecursionError:
            QMessageBox.warning(self,"Ошибка","Невозможно найти оптимальное решение...")
            return
        result = optimizer.getResult()
        document.generate(solution,result)
        document.save()
        if sys.platform == 'linux':
            subprocess.call(["xdg-open", filename]) 
        else:
            os.startfile(filename)
    def addC(self):
        if self.constraintCount==10:
            return
        self.constraintCount+=1
        self.constraintTable.setRowCount(self.constraintCount)
        for i in range(self.constraintCount):
            for j in range(self.targetCount+2):
                if self.constraintTable.item(i,j) == None:
                    self.constraintTable.setItem(i,j,QTableWidgetItem(str(0)))
        for i in range(self.constraintCount):
            lmBox = QComboBox()
            lmBox.addItems(['≥',"≤"])
            self.constraintTable.setCellWidget(i,self.targetCount,lmBox)
        self.constraintTable.resizeRowsToContents()
    def delC(self):
        if self.constraintCount==1:
            return
        self.constraintCount-=1
        self.constraintTable.setRowCount(self.constraintCount)
        for i in range(self.constraintCount):
            lmBox = QComboBox()
            lmBox.addItems(['≥',"≤"])
            self.constraintTable.setCellWidget(i,self.targetCount,lmBox)
        self.constraintTable.resizeRowsToContents()

    def addT(self):
        if self.targetCount==10:
            return
        self.constraintTable.removeColumn(self.targetCount)
        self.targetCount += 1
        self.targetTable.setColumnCount(self.targetCount)
        self.targetTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)])
        self.constraintTable.setColumnCount(self.targetCount+2)
        self.constraintTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)]+['']+["z"])
        for i in range(self.targetCount):
            self.targetTable.setItem(0,i,QTableWidgetItem(str(0))) 
        for i in range(self.targetCount+2):
            self.constraintTable.setItem(0,i,QTableWidgetItem(str(0)))
        self.constraintTable.resizeRowsToContents()
        lmBox = QComboBox()
        lmBox.addItems(['≥',"≤"])
        for i in range(self.constraintCount):
            self.constraintTable.setCellWidget(i,self.targetCount,lmBox)
    def delT(self):
        if self.targetCount == 2:
            return
        self.constraintTable.removeColumn(self.targetCount)
        self.targetCount -= 1
        self.targetTable.setColumnCount(self.targetCount)
        self.targetTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)])
        self.targetTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.constraintTable.setColumnCount(self.targetCount+2)
        self.constraintTable.setHorizontalHeaderLabels(["X"+str(i+1) for i in range(self.targetCount)]+['']+["z"])
        self.constraintTable.resizeRowsToContents()
        for i in range(self.targetCount):
            self.targetTable.setItem(0,i,QTableWidgetItem(str(0))) 
        for i in range(self.targetCount+2):
            self.constraintTable.setItem(0,i,QTableWidgetItem(str(0)))
        for i in range(self.constraintCount):
            lmBox = QComboBox()
            lmBox.addItems(['≥',"≤"])
            self.constraintTable.setCellWidget(i,self.targetCount,lmBox)
        self.targetTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())