# !/usr/bin/python3
# -*- coding: utf-8 -*-


# AuTomated LAbbook System (ATLAS)

import sys
import os
import docx
import sqlite3
import subprocess


from PyQt5.QtWidgets import QMainWindow, QApplication, QDesktopWidget, \
    QFrame, QLabel, QPushButton, QCheckBox, QAction, QFileDialog, QTableView, \
    QAbstractItemView, QTabWidget, QWidget, QLineEdit, QPlainTextEdit, \
    QCalendarWidget, QDateEdit, QComboBox

from PyQt5.QtGui import QFont, QStandardItemModel, QStandardItem, QPalette, QColor

from PyQt5.QtCore import Qt
#Key Global Variables

Experiments = {}

class Cal(QMainWindow):
    def __init__(self, parent=None):
        super(Cal, self).__init__(parent)
        self.setWindowTitle("Choose Section Date")
        self.resize(310,210)
        self.move(750,250)

        self.Cale = QCalendarWidget(self)
        self.Cale.setVerticalHeaderFormat(self.Cale.NoVerticalHeader)
        self.Cale.move(5,5)
        self.Cale.resize(300,200)

        self.Cale.clicked()


class Atlas(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.openFile = QAction('&Open', self)
        self.openFile.setShortcut('Ctrl+O')
        self.openFile.setStatusTip('Open new File')
        self.openFile.triggered.connect(self.showDialog)

        menubar = self.menuBar()
        self.fileMenu = menubar.addMenu('&File')
        self.fileMenu.addAction(self.openFile)

        self.resize(1200, 600)

        # Center Window
        qtRectangle = self.frameGeometry()
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        self.move(qtRectangle.topLeft())

        self.setWindowTitle('ATLAS')

        # StatusBar
        self.statusBar().showMessage('Ready')
        self.show()

        # Initialise Experiment Selection Frame
        self.ExpFrame = QFrame(self)
        print(self.ExpFrame.parentWidget())
        self.ExpFrame.move(5, 25)
        self.ExpFrame.resize(450, 200)
        self.ExpFrame.setFrameShape(QFrame.StyledPanel)
        self.ExpFrame.show()

        # Experiment Frame Label

        self.ExpLabel = QLabel(self.ExpFrame)
        self.ExpLabel.setText("Experiment Tables")
        self.ExpLabel.move(5, 1)
        newfont = QFont("Times", 8, QFont.Bold)
        self.ExpLabel.setFont(newfont)
        self.ExpLabel.show()

        # Experiment Table Generation Button

        self.ExpButton = QPushButton(self.ExpFrame)
        self.ExpButton.resize(120, 30)
        self.ExpButton.move(320, 20)
        self.ExpButton.setText("Generate")
        self.ExpButton.clicked.connect(self.GenExpList)
        self.ExpButton.show()

        # Experiment Check Buttons

        self.ExpCheck1 = QCheckBox(self.ExpFrame)
        self.ExpCheck1.move(15, 30)
        self.ExpCheck1.setText("Include Complete Experiments")
        self.ExpCheck1.show()

        self.ExpTable = QTableView(self.ExpFrame)
        self.ExpTable.move(10, 60)
        self.ExpTable.resize(430, 130)
        self.ExpTableModel = QStandardItemModel(self)
        self.ExpTable.setModel(self.ExpTableModel)
        self.ExpTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        row = []
        cell = QStandardItem("#")
        row.append(cell)
        cell = QStandardItem("Name")
        row.append(cell)
        self.ExpTable.horizontalHeader().hide()
        self.ExpTable.verticalHeader().hide()
        self.ExpTableModel.appendRow(row)
        self.ExpTable.horizontalHeader().setStretchLastSection(True)
        self.ExpTable.setColumnWidth(0,12)
        self.ExpTable.resizeRowsToContents()
        self.ExpTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        # self.ExpTable.clicked.connect(self.openExp)
        self.ExpTable.show()

        self.ModFrame = QFrame(self)
        self.ModFrame.move(5, 230)
        self.ModFrame.resize(450, 350)
        self.ModFrame.setFrameShape(QFrame.StyledPanel)
        self.ModFrame.show()

        self.ModLabel = QLabel(self.ModFrame)
        self.ModLabel.setText("Section Viewer")
        self.ModLabel.move(5, 1)
        newfont = QFont("Times", 8, QFont.Bold)
        self.ModLabel.setFont(newfont)
        self.ModLabel.show()

        self.ModTable = QTableView(self.ModFrame)
        self.ModTable.move(10, 20)
        self.ModTable.resize(430, 320)
        self.ModTableModel = QStandardItemModel(self)
        self.ModTable.setModel(self.ModTableModel)
        self.ModTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        row = []
        cell = QStandardItem("Date")
        row.append(cell)
        cell = QStandardItem("Name")
        row.append(cell)
        self.ModTable.horizontalHeader().hide()
        self.ModTable.verticalHeader().hide()
        self.ModTableModel.appendRow(row)
        self.ModTable.horizontalHeader().setStretchLastSection(True)
        self.ModTable.setColumnWidth(0,80)
        self.ModTable.resizeRowsToContents()
        self.ModTable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ModTable.clicked.connect(self.openExp)
        self.ModTable.show()


        # self.DetFrame = QFrame(self)
        # self.DetFrame.move(460, 25)
        # self.DetFrame.resize(735, 555)
        # self.DetFrame.setFrameShape(QFrame.StyledPanel)
        # self.DetFrame.show()
        #
        # self.DetLabel = QLabel(self.DetFrame)
        # self.DetLabel.setText("Detailed View")
        # self.DetLabel.move(5, 1)
        # newfont = QFont("Times", 8, QFont.Bold)
        # self.DetLabel.setFont(newfont)
        # self.DetLabel.show()

        self.DetTabs = QTabWidget(self)
        self.DetTab1 = QWidget(self)
        self.DetTab1.setAutoFillBackground(True)
        self.DetTab2 = QWidget(self)
        self.DetTab2.setAutoFillBackground(True)
        self.DetTab3 = QWidget(self)
        self.DetTab3.setAutoFillBackground(True)
        self.DetTabs.move(460, 25)
        self.DetTabs.resize(735, 556)
        self.DetTabs.addTab(self.DetTab1,"DetTab1")
        self.DetTabs.addTab(self.DetTab2,"DetTab2")
        self.DetTabs.addTab(self.DetTab3, "New Template")
        self.DetTabs.show()

        self.DetTabs.currentChanged.connect(self.TabChange)

        #Load variable list when tab is activated
        # self.DetTab3.actionEvent()

        self.DTNew_Lab_Title = QLabel(self.DetTab3)
        self.DTNew_Lab_Title.setText("Title")
        self.DTNew_Lab_Title.move(5, 2)
        newfont = QFont("Times", 8, QFont.Bold)
        self.DTNew_Lab_Title.setFont(newfont)
        self.DTNew_Lab_Title.show()

        self.DTNew_Text_Title = QLineEdit(self.DetTab3)
        self.DTNew_Text_Title.setText("Template Title")
        self.DTNew_Text_Title.move(3,23)
        self.DTNew_Text_Title.resize(723,17)
        self.DTNew_Text_Title.show()

        self.DTNew_Lab_Date = QLabel(self.DetTab3)
        self.DTNew_Lab_Date.setText("Date")
        self.DTNew_Lab_Date.move(5, 47)
        newfont = QFont("Times", 8, QFont.Bold)
        self.DTNew_Lab_Date.setFont(newfont)
        self.DTNew_Lab_Date.show()

        self.DTNew_But_Date = QDateEdit(self.DetTab3)
        self.DTNew_But_Date.move(3,68)
        self.DTNew_But_Date.resize(80,19)
        self.DTNew_But_Date.setCalendarPopup(True)
        self.DTNew_But_Date.show()

        self.DTNew_Lab_Section = QLabel(self.DetTab3)
        self.DTNew_Lab_Section.setText("Section Text")
        self.DTNew_Lab_Section.move(5, 94)
        newfont = QFont("Times", 8, QFont.Bold)
        self.DTNew_Lab_Section.setFont(newfont)
        self.DTNew_Lab_Section.show()

        self.DTNew_Text_Section = QPlainTextEdit(self.DetTab3)
        self.DTNew_Text_Section.appendPlainText("Section Text")
        self.DTNew_Text_Section.move(3,116)
        self.DTNew_Text_Section.resize(723,200)
        self.DTNew_Text_Section.show()

        self.DTNew_Lab_Var = QLabel(self.DetTab3)
        self.DTNew_Lab_Var.setText("Insert Variable")
        self.DTNew_Lab_Var.move(5, 326)
        newfont = QFont("Times", 8, QFont.Bold)
        self.DTNew_Lab_Var.setFont(newfont)
        self.DTNew_Lab_Var.show()

        self.DTNew_Text_Section = QComboBox(self.DetTab3)
        # self.DTNew_Text_Section.appendPlainText("Section Text")
        self.DTNew_Text_Section.move(3,346)
        self.DTNew_Text_Section.show()

        #Check Variable Database
        if os.path.exists("./Variables/Variables.sqlite3") == False:
            #Create Variables Database
            conn = sqlite3.connect("./Variables/Variables.sqlite3")
            c = conn.cursor()

            #Create Antibiotics Table
            c.execute("CREATE TABLE Antibiotics(id INTEGER PRIMARY KEY, name TEXT)")
            #Add Standard Antibiotics
            #ant="Carbenicillin"
            #c.execute("INSERT INTO Antibiotics(name) VALUES (?)",(ant,))
            ant=["Carbenicillin"]
            c.execute("INSERT INTO Antibiotics(name) VALUES (?)",(ant,))
            c.execute("SELECT * FROM Antibiotics")
            print(c.fetchall())

            conn.commit()
            conn.close()



    def TabChange(self, i):
        if i == 2:
            print("yes")


    def openExp(self,clickedIndex):
        Exp = clickedIndex.sibling(clickedIndex.row(),0).data()
        print(Exp)
        os.startfile(Experiments[Exp])

    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', '.')
        print(fname[0])

    def GenExpList(self):
        self.ExpTableModel.setRowCount(0)
        for root, dirs, files in os.walk(".\Experiments"):
            for file in files:
                if file.startswith("Experiment"):
                    if "." in file.split()[1]:
                        doc = docx.Document(root + "\\" + file)
                        if doc.core_properties.subject != "Complete" or self.ExpCheck1.isChecked():
                            row = []
                            row.append(QStandardItem(file.split()[1][:file.split()[1].index(".")]))
                            row.append(QStandardItem(doc.core_properties.title))
                            self.ExpTableModel.appendRow(row)
                            self.ExpTable.resizeRowsToContents()
                            Experiments[file.split()[1][:file.split()[1].index(".")]] = root+"\\"+file

                    else:
                        doc = docx.Document(root+"\\"+file)
                        if doc.core_properties.subject != "Complete" or self.ExpCheck1.isChecked():
                            row = []
                            row.append(QStandardItem(file.split()[1]))
                            row.append(QStandardItem(doc.core_properties.title))
                            self.ExpTableModel.appendRow(row)
                            self.ExpTable.resizeRowsToContents()
                            Experiments[file.split()[1]] = root+"\\"+file
        doc = docx.Document(os.path.realpath(".\demo 1.docx"))
        print(doc.paragraphs[1].text)
        print(Experiments)

    def OpenCal(self):
        print("hello")
        self.DTNew_Cal.show()


# class DatePicker(QCalendarWidget):
#     def __init__(self, parent=None):


        # if fname[0]:
        #     f = open(fname[0], 'r')
        #
        #     with f:
        #         data = f.read()
        #         self.textEdit.setText(data)

# Is this file being run directly?
if __name__ == '__main__':
    app = QApplication(sys.argv)

    #
    start = Atlas()
    #

    sys.exit(app.exec_())
