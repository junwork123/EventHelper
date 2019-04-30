# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'P:\Github\EventHelper\ui\main.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QStringListModel, pyqtSignal, QThread
from PyQt5.QtWidgets import QCompleter, QDialog

import time, threading
from datetime import datetime
from excelGenerator import generateExcel
from Event import *

class Ui_Dialog(QDialog, QThread):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.setEnabled(True)
        Dialog.resize(369, 445)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(190, 400, 161, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.tabWidget = QtWidgets.QTabWidget(Dialog)
        self.tabWidget.setGeometry(QtCore.QRect(20, 20, 331, 341))
        self.tabWidget.setObjectName("tabWidget")
        self.eventTab = QtWidgets.QWidget()
        self.eventTab.setObjectName("tab")
        self.groupBox = QtWidgets.QGroupBox(self.eventTab)
        self.groupBox.setGeometry(QtCore.QRect(10, 5, 301, 80))
        self.groupBox.setObjectName("groupBox")

        now = datetime.now()
        self.startTime = QtWidgets.QDateTimeEdit(self.groupBox)
        self.startTime.setGeometry(QtCore.QRect(10, 20, 194, 22))
        self.startTime.setDate(QtCore.QDate(now.year, now.month, now.day))
        self.startTime.setCalendarPopup(True)
        self.startTime.setObjectName("startTime")
        self.endTime = QtWidgets.QDateTimeEdit(self.groupBox)
        self.endTime.setGeometry(QtCore.QRect(10, 50, 194, 22))
        self.endTime.setDate(QtCore.QDate(now.year, now.month, now.day))
        self.endTime.setCalendarPopup(True)
        self.endTime.setObjectName("endTime")
        self.groupBox_2 = QtWidgets.QGroupBox(self.eventTab)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 90, 301, 221))
        self.groupBox_2.setObjectName("groupBox_2")

        self.checkBox_Cherry = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox_Cherry.setGeometry(QtCore.QRect(10, 20, 151, 20))
        self.checkBox_Cherry.setObjectName("checkBox_Cherry")
        self.checkBox_URL = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox_URL.setGeometry(QtCore.QRect(10, 50, 101, 16))
        self.checkBox_URL.setObjectName("checkBox_URL")
        self.checkBox_Keyword = QtWidgets.QCheckBox(self.groupBox_2)
        self.checkBox_Keyword.setGeometry(QtCore.QRect(10, 80, 101, 16))
        self.checkBox_Keyword.setObjectName("checkBox_Keyword")

        self.frame = QtWidgets.QFrame(self.groupBox_2)
        self.frame.setGeometry(QtCore.QRect(20, 100, 261, 111))
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.keywordEdit1 = QtWidgets.QLineEdit(self.frame)
        self.keywordEdit1.setGeometry(QtCore.QRect(65, 6, 181, 20))
        self.keywordEdit1.setObjectName("keywordEdit1")
        self.keywordEdit2 = QtWidgets.QLineEdit(self.frame)
        self.keywordEdit2.setGeometry(QtCore.QRect(65, 31, 181, 20))
        self.keywordEdit2.setObjectName("keywordEdit2")
        self.keywordEdit3 = QtWidgets.QLineEdit(self.frame)
        self.keywordEdit3.setGeometry(QtCore.QRect(65, 56, 181, 20))
        self.keywordEdit3.setObjectName("keywordEdit3")
        self.keywordEdit4 = QtWidgets.QLineEdit(self.frame)
        self.keywordEdit4.setGeometry(QtCore.QRect(65, 81, 181, 20))
        self.keywordEdit4.setObjectName("keywordEdit4")
        self.keywordLabel1 = QtWidgets.QLabel(self.frame)
        self.keywordLabel1.setGeometry(QtCore.QRect(10, 10, 56, 12))
        self.keywordLabel1.setObjectName("keywordLabel1")
        self.keywordLabel2 = QtWidgets.QLabel(self.frame)
        self.keywordLabel2.setGeometry(QtCore.QRect(9, 35, 56, 12))
        self.keywordLabel2.setObjectName("keywordLabel2")
        self.keywordLabel3 = QtWidgets.QLabel(self.frame)
        self.keywordLabel3.setGeometry(QtCore.QRect(9, 60, 56, 12))
        self.keywordLabel3.setObjectName("keywordLabel3")
        self.keywordLabel4 = QtWidgets.QLabel(self.frame)
        self.keywordLabel4.setGeometry(QtCore.QRect(9, 85, 56, 12))
        self.keywordLabel4.setObjectName("keywordLabel4")

        self.tabWidget.addTab(self.eventTab, "")
        self.prizeTab = QtWidgets.QWidget()
        self.prizeTab.setObjectName("prizeTab")
        self.prizeLabel = QtWidgets.QLabel(self.prizeTab)
        self.prizeLabel.setGeometry(QtCore.QRect(16, 15, 56, 12))
        self.prizeLabel.setObjectName("prizeLabel")
        self.prizeComboBox = QtWidgets.QComboBox(self.prizeTab)
        self.prizeComboBox.setGeometry(QtCore.QRect(80, 10, 76, 22))
        self.prizeComboBox.setObjectName("prizeComboBox")
        # 상품종류는 1~9개
        for i in range(1,10):
            self.prizeComboBox.addItem(str(i))

        self.prizeEdit = []
        for i in range(0,9):
            self.prizeEdit.append(QtWidgets.QLineEdit(self.prizeTab))
            self.prizeEdit[i].setObjectName("prizeEdit" + str(i))
            self.prizeEdit[i].hide()

        self.prizeEdit[0].show()
        self.prizeEdit[0].setGeometry(QtCore.QRect(10, 40, 201, 20))
        self.prizeEdit[1].setGeometry(QtCore.QRect(10, 70, 201, 20))
        self.prizeEdit[2].setGeometry(QtCore.QRect(10, 100, 201, 20))
        self.prizeEdit[3].setGeometry(QtCore.QRect(10, 130, 201, 20))
        self.prizeEdit[4].setGeometry(QtCore.QRect(10, 160, 201, 20))
        self.prizeEdit[5].setGeometry(QtCore.QRect(10, 190, 201, 20))
        self.prizeEdit[6].setGeometry(QtCore.QRect(10, 220, 201, 20))
        self.prizeEdit[7].setGeometry(QtCore.QRect(10, 250, 201, 20))
        self.prizeEdit[8].setGeometry(QtCore.QRect(10, 280, 201, 20))

        self.prizeSpinBox = []
        for i in range(0,9):
            self.prizeSpinBox.append(QtWidgets.QSpinBox(self.prizeTab))
            self.prizeSpinBox[i].setObjectName("prizeSpinBox" + str(i))
            self.prizeSpinBox[i].hide()

        self.prizeSpinBox[0].show()
        self.prizeSpinBox[0].setGeometry(QtCore.QRect(230, 40, 42, 22))
        self.prizeSpinBox[1].setGeometry(QtCore.QRect(230, 70, 42, 22))
        self.prizeSpinBox[2].setGeometry(QtCore.QRect(230, 100, 42, 22))
        self.prizeSpinBox[3].setGeometry(QtCore.QRect(230, 130, 42, 22))
        self.prizeSpinBox[4].setGeometry(QtCore.QRect(230, 160, 42, 22))
        self.prizeSpinBox[5].setGeometry(QtCore.QRect(230, 190, 42, 22))
        self.prizeSpinBox[6].setGeometry(QtCore.QRect(230, 220, 42, 22))
        self.prizeSpinBox[7].setGeometry(QtCore.QRect(230, 250, 42, 22))
        self.prizeSpinBox[8].setGeometry(QtCore.QRect(230, 280, 42, 22))
        self.tabWidget.addTab(self.prizeTab, "")

        self.progressBar = QtWidgets.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(20, 370, 331, 23))
        self.progressBar.setMaximum(100)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progressBar.setObjectName("progressBar")
        self.calcProgressBar = ExternalThread()
        self.calcProgressBar.notifyProgress.connect(self.onProgress)
        #self.calcProgressBar.start()
        self.t = threading.Thread(target=self.proceed)

        self.retranslateUi(Dialog)
        self.tabWidget.setCurrentIndex(0)
        self.buttonBox.accepted.connect(self.onButtonCliked)
        self.buttonBox.rejected.connect(Dialog.reject)
        self.checkBox_URL.clicked['bool'].connect(self.checkBox_URL.setChecked)
        self.checkBox_Keyword.clicked['bool'].connect(self.checkBox_Keyword.setChecked)
        self.checkBox_Cherry.clicked['bool'].connect(self.checkBox_Cherry.setChecked)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.groupBox.setTitle(_translate("Dialog", "이벤트 기간"))
        self.groupBox_2.setTitle(_translate("Dialog", "당첨자 조건"))
        self.checkBox_Cherry.setText(_translate("Dialog", "체리피커 필터링"))
        self.checkBox_Cherry.setChecked(True)
        self.checkBox_URL.setText(_translate("Dialog", "URL 첨부"))
        self.checkBox_URL.setChecked(True)
        self.checkBox_Keyword.setText(_translate("Dialog", "키워드 포함"))
        self.checkBox_Keyword.setChecked(True)
        self.keywordLabel1.setText(_translate("Dialog", "키워드 1"))
        self.keywordLabel2.setText(_translate("Dialog", "키워드 2"))
        self.keywordLabel3.setText(_translate("Dialog", "키워드 3"))
        self.keywordLabel4.setText(_translate("Dialog", "키워드 4"))

        self.endTime.dateTimeChanged.connect(self.onOverStartTime)


        #자동완성기능(추후 제작예정)
        prizeModel = QStringListModel()
        prizeModel.setStringList([])
        completer = QCompleter()
        completer.setModel(prizeModel)
        #self.prizeEdit1.setCompleter(completer)


        self.tabWidget.setTabText(self.tabWidget.indexOf(self.eventTab), _translate("Dialog", "이벤트 설정"))
        self.prizeLabel.setText(_translate("Dialog", "상품 종류"))
        # 콤보박스에 이벤트슬롯 추가
        self.prizeComboBox.currentTextChanged.connect(self.onSelected)
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.prizeTab), _translate("Dialog", "경품 설정"))

    def run(self):
        once = True
        if(once is True):
            self.proceed()
            once = False

    def onButtonCliked(self):
        self.t.start()
        self.calcProgressBar.start()


    # 상품종류(콤보박스) 값에 따라 텍스트에디트, 스핀박스 등장
    def onSelected(self):
        selectedNum = int(self.prizeComboBox.currentText())
        for i in range(0, 9):
            if(i < selectedNum):
                self.prizeEdit[i].show()
                self.prizeSpinBox[i].show()
            else:
                self.prizeEdit[i].hide()
                self.prizeSpinBox[i].hide()

    # 종료시간이 시작시간보다 빠를 경우 메시지 발생
    def onOverStartTime(self):
        print(self.startTime.date())
        print(self.endTime.date())
        if(self.startTime.date() > self.endTime.date()):
            popup = QtWidgets.QMessageBox(self)
            popup.setGeometry(100, 200, 100, 100)
            popup.information(self,"Error", "종료날짜가 시작날짜보다 빠릅니다.",QtWidgets.QMessageBox.Yes)

    def onProgress(self, i):
        self.progressBar.setValue(i)
        if(self.progressBar.value() > self.progressBar.maximum()):
            self.close()

    def proceed(self):
        # 결과를 제출할 이벤트 인스턴스 생성
        myEvent = Event()
        myEvent.setStartTime(self.startTime.text())
        myEvent.setEndTime(self.endTime.text())

        # 상품을 사전에 등록
        dic = {}
        selectedNum = int(self.prizeComboBox.currentText())
        for i in range(0,selectedNum):
            dic[self.prizeEdit[i].text()] = self.prizeSpinBox[i].value()

        myEvent.setPrize(dic)
        print("당첨자 엑셀을 생성합니다")

        # 체리피커 필터링
        if (self.checkBox_Cherry.isChecked() is True):
            myEvent.setOpt_Cherry(["초대", "함께해", "함께해요"])
            print("체리피커 필터링이 활성화 되었습니다")

        # URL 체크
        if(self.checkBox_URL.isChecked() is True):
            myEvent.setOpt_URL(["www", "facebook"])
            print("URL 체크가 활성화 되었습니다")

        # 이벤트 키워드 체크
        if (self.checkBox_Keyword.isChecked() is True):
            if not (self.keywordEdit1.text() is None):
                myEvent.setOpt_Keywords([self.keywordEdit1.text()])
            if not (self.keywordEdit2.text() is None):
                myEvent.setOpt_Keywords([self.keywordEdit2.text()])
            if not (self.keywordEdit3.text() is None):
                myEvent.setOpt_Keywords([self.keywordEdit3.text()])
            if not (self.keywordEdit4.text() is None):
                myEvent.setOpt_Keywords([self.keywordEdit4.text()])
            print("키워드 체크가 활성화 되었습니다")

        print("체리필터 키워드 : " + str(myEvent.getOpt_Cherry()))
        print("URL 체크 키워드 : " + str(myEvent.getOpt_URL()))
        print("이벤트 당첨 키워드 : " + str(myEvent.getOpt_Keywords()))

        generateExcel(myEvent)
        self.progressBar.setValue(self.progressBar.maximum())
        popup = QtWidgets.QMessageBox(self)
        popup.setGeometry(100, 200, 100, 100)
        popup.information(self, "Complete", "당첨자 엑셀이 생성되었습니다.", QtWidgets.QMessageBox.Yes)

class ExternalThread(QThread):
    notifyProgress = pyqtSignal(int)

    def run(self):

        for i in range(0,50):
            self.notifyProgress.emit(i)
            time.sleep(0.5)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

