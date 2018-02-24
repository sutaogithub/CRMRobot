import sys
from PyQt5.QtGui import QMovie
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel
from PyQt5 import QtCore, QtWidgets


class LoadingDialog(QDialog):

    def __init__(self, parent=None):
        super(LoadingDialog, self).__init__(parent)
        self.initUI()
        self.setModal(True)
        self.resize(150, 150)

    def initUI(self):
        layout = QVBoxLayout()
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        # self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.loadingLabel = QLabel()
        self.loadingLabel.setScaledContents(True)
        movie = QMovie("loading.gif")
        self.loadingLabel.setMovie(movie)
        movie.start()
        self.waitingLabel = QLabel()
        self.waitingLabel.setText("正在执行，请稍后...")
        layout.addWidget(self.loadingLabel, alignment=QtCore.Qt.AlignHCenter)
        layout.addWidget(self.waitingLabel, alignment=QtCore.Qt.AlignHCenter)
        self.setLayout(layout)



