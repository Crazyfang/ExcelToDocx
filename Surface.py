# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Surface.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(549, 555)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox_Operate = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_Operate.setGeometry(QtCore.QRect(10, 10, 531, 211))
        self.groupBox_Operate.setObjectName("groupBox_Operate")
        self.label_FileOrDirectory = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_FileOrDirectory.setGeometry(QtCore.QRect(20, 50, 121, 16))
        self.label_FileOrDirectory.setObjectName("label_FileOrDirectory")
        self.lineEdit_FileOrDirectory = QtWidgets.QLineEdit(self.groupBox_Operate)
        self.lineEdit_FileOrDirectory.setGeometry(QtCore.QRect(150, 50, 241, 21))
        self.lineEdit_FileOrDirectory.setText("")
        self.lineEdit_FileOrDirectory.setObjectName("lineEdit_FileOrDirectory")
        self.pushButton_FileOrDirectory = QtWidgets.QPushButton(self.groupBox_Operate)
        self.pushButton_FileOrDirectory.setGeometry(QtCore.QRect(410, 41, 91, 41))
        self.pushButton_FileOrDirectory.setAcceptDrops(False)
        self.pushButton_FileOrDirectory.setAutoFillBackground(False)
        self.pushButton_FileOrDirectory.setFlat(False)
        self.pushButton_FileOrDirectory.setObjectName("pushButton_FileOrDirectory")
        self.progressBar = QtWidgets.QProgressBar(self.groupBox_Operate)
        self.progressBar.setGeometry(QtCore.QRect(100, 142, 291, 31))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.label_Schedule = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_Schedule.setGeometry(QtCore.QRect(20, 150, 60, 16))
        self.label_Schedule.setObjectName("label_Schedule")
        self.pushButton_Start = QtWidgets.QPushButton(self.groupBox_Operate)
        self.pushButton_Start.setGeometry(QtCore.QRect(410, 140, 91, 41))
        self.pushButton_Start.setObjectName("pushButton_Start")
        self.label_SelectType = QtWidgets.QLabel(self.groupBox_Operate)
        self.label_SelectType.setGeometry(QtCore.QRect(20, 100, 91, 16))
        self.label_SelectType.setObjectName("label_SelectType")
        self.radioButton_File = QtWidgets.QRadioButton(self.groupBox_Operate)
        self.radioButton_File.setGeometry(QtCore.QRect(140, 100, 100, 20))
        self.radioButton_File.setChecked(True)
        self.radioButton_File.setObjectName("radioButton_File")
        self.radioButton_Directory = QtWidgets.QRadioButton(self.groupBox_Operate)
        self.radioButton_Directory.setGeometry(QtCore.QRect(270, 100, 100, 20))
        self.radioButton_Directory.setObjectName("radioButton_Directory")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(10, 230, 531, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.groupBox_Adjust = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_Adjust.setGeometry(QtCore.QRect(10, 250, 531, 301))
        self.groupBox_Adjust.setObjectName("groupBox_Adjust")
        self.listWidget_Adjust = QtWidgets.QListWidget(self.groupBox_Adjust)
        self.listWidget_Adjust.setGeometry(QtCore.QRect(10, 30, 441, 261))
        self.listWidget_Adjust.setDragEnabled(True)
        self.listWidget_Adjust.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.listWidget_Adjust.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.listWidget_Adjust.setObjectName("listWidget_Adjust")
        self.pushButton_Combine = QtWidgets.QPushButton(self.groupBox_Adjust)
        self.pushButton_Combine.setGeometry(QtCore.QRect(450, 111, 81, 71))
        self.pushButton_Combine.setObjectName("pushButton_Combine")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "信贷过程优化"))
        self.groupBox_Operate.setTitle(_translate("MainWindow", "操作区"))
        self.label_FileOrDirectory.setText(_translate("MainWindow", "选择文件或者文件夹:"))
        self.pushButton_FileOrDirectory.setText(_translate("MainWindow", "选择..."))
        self.label_Schedule.setText(_translate("MainWindow", "处理进度:"))
        self.pushButton_Start.setText(_translate("MainWindow", "开始处理"))
        self.label_SelectType.setText(_translate("MainWindow", "切换选取方式:"))
        self.radioButton_File.setText(_translate("MainWindow", "文件"))
        self.radioButton_Directory.setText(_translate("MainWindow", "文件夹"))
        self.groupBox_Adjust.setTitle(_translate("MainWindow", "顺序调整"))
        self.pushButton_Combine.setText(_translate("MainWindow", "合并"))
