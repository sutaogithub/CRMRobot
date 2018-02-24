# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'page.ui'
#
# Created by: PyQt5 UI code generator 5.10
#
# WARNING! All changes made in this file will be lost!
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, QThread
from PyQt5.QtWidgets import QFileDialog, QLineEdit, QLabel, QPushButton, QSpinBox, QDoubleSpinBox, QMenuBar, QStatusBar, \
    QMessageBox
from selenium.common.exceptions import UnexpectedAlertPresentException

from loading import LoadingDialog
from auto import *
import threading


class Signal(QThread):
    msgSignal = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(Signal, self).__init__(parent)
        self.msg = None

    def run(self):
        self.msgSignal.emit(self.msg)

    def emit(self, msg):
        self.msg = msg
        self.start()

    def connect(self, slot):
        self.msgSignal.connect(slot)


class Ui_MainWindow(QtWidgets.QMainWindow):

    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.crmAuto = None
        self.setupUi()
        self.loading = LoadingDialog(self)
        self.backThread = None
        self.messageDialog = None
        self.signal = Signal()
        self.signal.connect(self.printMessage)
        self.logFile = open("log", "a")

    def setupUi(self):
        self.resize(572, 387)
        self.setObjectName("centralwidget")
        self.label = QLabel(self)
        self.label.setGeometry(QtCore.QRect(30, 40, 54, 12))
        self.label.setObjectName("label")
        self.label_2 = QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(30, 80, 54, 12))
        self.label_2.setObjectName("label_2")
        self.excelPath = QLineEdit(self)
        self.excelPath.setGeometry(QtCore.QRect(90, 30, 381, 31))
        self.excelPath.setObjectName("excelPath")
        self.iniPath = QLineEdit(self)
        self.iniPath.setGeometry(QtCore.QRect(90, 70, 381, 31))
        self.iniPath.setObjectName("iniPath")

        self.label_3 = QLabel(self)
        self.label_3.setGeometry(QtCore.QRect(40, 130, 54, 12))
        self.label_3.setObjectName("label_3")
        self.label_4 = QLabel(self)
        self.label_4.setGeometry(QtCore.QRect(50, 170, 54, 12))
        self.label_4.setObjectName("label_4")
        self.account = QLineEdit(self)
        self.account.setGeometry(QtCore.QRect(90, 120, 381, 31))
        self.account.setObjectName("account")
        self.password = QLineEdit(self)
        self.password.setGeometry(QtCore.QRect(90, 160, 381, 31))
        self.password.setObjectName("password")
        self.password.setEchoMode(QLineEdit.Password)
        self.label_5 = QLabel(self)
        self.label_5.setGeometry(QtCore.QRect(3, 230, 80, 12))
        self.label_5.setObjectName("label_5")
        self.label_6 = QLabel(self)
        self.label_6.setGeometry(QtCore.QRect(210, 230, 54, 12))
        self.label_6.setObjectName("label_6")
        self.rowBox = QSpinBox(self)
        self.rowBox.setMinimum(3)
        self.rowBox.setGeometry(QtCore.QRect(270, 220, 41, 31))
        self.rowBox.setObjectName("spinBox")
        self.intervalBox = QDoubleSpinBox(self)
        self.intervalBox.setMinimum(0.05)
        self.intervalBox.setValue(1)
        self.intervalBox.setGeometry(QtCore.QRect(90, 220, 62, 31))
        self.intervalBox.setObjectName("doubleSpinBox")
        self.pushButton = QPushButton(self)
        self.pushButton.setGeometry(QtCore.QRect(480, 30, 75, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QPushButton(self)
        self.pushButton_2.setGeometry(QtCore.QRect(480, 70, 75, 23))
        self.pushButton_2.setObjectName("pushButton_2")

        self.pushButton_3 = QPushButton(self)
        self.pushButton_3.setGeometry(QtCore.QRect(20, 280, 91, 41))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QPushButton(self)
        self.pushButton_4.setGeometry(QtCore.QRect(150, 280, 91, 41))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QPushButton(self)
        self.pushButton_5.setGeometry(QtCore.QRect(280, 280, 91, 41))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_6 = QPushButton(self)
        self.pushButton_6.setGeometry(QtCore.QRect(410, 280, 125, 41))
        self.pushButton_6.setObjectName("pushButton_6")

        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 572, 23))
        self.menubar.setObjectName("menubar")
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi()
        self.pushButton.clicked.connect(self.getExcelFilePath)
        self.pushButton_2.clicked.connect(self.getIniFilePath)
        self.pushButton_3.clicked.connect(self.login)
        self.pushButton_4.clicked.connect(self.startInput)
        self.pushButton_5.clicked.connect(self.cancelAlert)
        self.pushButton_6.clicked.connect(self.refreshFile)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "CRM+辅助录入脚本"))
        self.label.setText(_translate("MainWindow", "Excel文件:"))
        self.label_2.setText(_translate("MainWindow", "配置文件:"))
        self.pushButton.setText(_translate("MainWindow", "选择文件"))
        self.pushButton_2.setText(_translate("MainWindow", "选择文件"))
        self.label_3.setText(_translate("MainWindow", "用户名"))
        self.label_4.setText(_translate("MainWindow", "密码"))
        self.label_5.setText(_translate("MainWindow", "录入间隔（秒）"))
        self.label_6.setText(_translate("MainWindow", "录入行号"))
        self.pushButton_3.setText(_translate("MainWindow", "登陆"))
        self.pushButton_4.setText(_translate("MainWindow", "开始录入"))
        self.pushButton_5.setText(_translate("MainWindow", "去掉弹窗"))
        self.pushButton_6.setText(_translate("MainWindow", "刷新excel和配置文件"))

    def refreshFile(self):
        if self.crmAuto == None:
            self.printMessage("请先登录")
            return
        excelFile = str(self.excelPath.text())
        iniFile = str(self.iniPath.text())
        if excelFile == "":
            self.printMessage("请选择excel目录")
            return
        if iniFile == "":
            self.printMessage("请选择配置文件目录")
            return
        self.crmAuto.set_excel(excelFile)
        self.crmAuto.set_xpath(iniFile)
        self.printMessage("文件已刷新")

    def getExcelFilePath(self):
        filename = str(QFileDialog.getOpenFileName(self, filter="Excel File (*.xls *xlsx)"))
        list = filename.split("'")
        self.excelPath.setText(list[1])

    def getIniFilePath(self):
        filename = str(QFileDialog.getOpenFileName(self, filter="INI File (*.ini)"))
        list = filename.split("'")
        self.iniPath.setText(list[1])

    def login(self):
        excelFile = str(self.excelPath.text())
        iniFile = str(self.iniPath.text())
        account = str(self.account.text())
        password = str(self.password.text())
        if excelFile == "":
            self.printMessage("请选择excel目录")
            return
        if iniFile == "":
            self.printMessage("请选择配置文件目录")
            return
        if account == "":
            self.printMessage("请输入账号")
            return
        if password == "":
            self.printMessage("请输入密码")
            return
        self.crmAuto = CRMAuto(excelFile, iniFile)
        self.backThread = threading.Thread(target=self.onlogin, args=(account, password))
        self.loading.show()
        self.backThread.start()

    def onlogin(self, account, password):
        try:
            self.crmAuto.login(account, password)
        except UnexpectedAlertPresentException as e:
            self.signal.emit("出现严重错误，请重新登录再试\n" + traceback.format_exc())
        except Exception as e:
            traceback.print_exc(file=self.logFile)
            self.signal.emit("出错了,请重试\n" + traceback.format_exc())
        finally:
            self.loading.close()

    def startInput(self):
        if self.crmAuto == None:
            self.printMessage("请先登录")
            return
        row = self.rowBox.value()
        interval = self.intervalBox.value()
        self.crmAuto.set_row_index(row)
        self.crmAuto.set_interval(interval)
        self.printMessage("5秒后开始录入，请切换到浏览器窗口")
        self.loading.show()
        self.backThread = threading.Thread(target=self.autoInput)
        self.backThread.start()

    def autoInput(self):
        try:
            self.crmAuto.auto_input()
        except UnexpectedAlertPresentException as e:
            self.signal.emit("出现严重错误，请重新登录再试\n" + traceback.format_exc())
        except Exception as e:
            traceback.print_exc(file=self.logFile)
            self.signal.emit("出错了,请重试\n" + traceback.format_exc())
        finally:
            self.loading.close()

    def printMessage(self, message):
        self.messageDialog = QMessageBox.information(self, "警告", message)

    def cancelAlert(self):
        if self.crmAuto == None:
            self.printMessage("请先登录")
            return
        self.crmAuto.close_alert()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = Ui_MainWindow()
    ex.show()
    app.exec_()
