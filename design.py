from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox, QMainWindow, QAction
from PyQt5.QtGui import QIcon
import json



class Ui_Main(QtWidgets.QWidget):
    def setupUi(self, Main):
        Main.setObjectName("Main")
        Main.resize(800, 480)



        self.QtStack = QtWidgets.QStackedLayout()



        self.stack1 = QtWidgets.QWidget() # главная страница
        self.stack2 = QtWidgets.QWidget() # исходные данные
        self.stack3 = QtWidgets.QWidget() # паспорт АФС 1
        self.stack4 = QtWidgets.QWidget() # паспорт АФС 2
        self.stack5 = QtWidgets.QWidget() # паспорт АФС 3
        self.stack6 = QtWidgets.QWidget() # геодезия
        self.stack7 = QtWidgets.QWidget() # примечания
        self.stack8 = QtWidgets.QWidget() # журнал АФС



        self.Window1UI() # главная страница
        self.Window2UI() # исходные данные
        self.Window3UI() # паспорт АФС 1
        self.Window4UI() # паспорт АФС 2
        self.Window5UI() # паспорт АФС 3
        self.Window6UI() # геодезия
        self.Window7UI() # геодезия
        self.Window8UI() # журнал АФС


        self.QtStack.addWidget(self.stack1)
        self.QtStack.addWidget(self.stack2)
        self.QtStack.addWidget(self.stack3)
        self.QtStack.addWidget(self.stack4)
        self.QtStack.addWidget(self.stack5)
        self.QtStack.addWidget(self.stack6)
        self.QtStack.addWidget(self.stack7)
        self.QtStack.addWidget(self.stack8)





####################################################################################################
##########
#                                  главная страница
##########
#####################################################################################################


    def Window1UI(self):
        self.stack1.setObjectName("MainWindow")
        self.stack1.resize(1200, 800)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack1.sizePolicy().hasHeightForWidth())
        self.stack1.setSizePolicy(sizePolicy)
        self.stack1.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack1.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack1.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack1)
        self.centralwidget.setObjectName("centralwidget")
        self.stack1.setWindowIcon(QIcon("images/logo.png"))
        
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
    
        self.pushButton1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton1.setGeometry(QtCore.QRect(355, 260, 489, 71))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton1.setFont(font)
        self.pushButton1.setMouseTracking(True)
        self.pushButton1.setAutoFillBackground(False)
        self.pushButton1.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton1.setAutoRepeat(False)
        self.pushButton1.setObjectName("pushButton1")
        
        self.pushButton2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton2.setGeometry(QtCore.QRect(355, 365, 489, 71))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton2.setFont(font)
        self.pushButton2.setMouseTracking(True)
        self.pushButton2.setAutoFillBackground(False)
        self.pushButton2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton2.setAutoRepeat(False)
        self.pushButton2.setObjectName("pushButton2")
        
        self.pushButton3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton3.setGeometry(QtCore.QRect(355, 470, 489, 71))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton3.setFont(font)
        self.pushButton3.setMouseTracking(True)
        self.pushButton3.setAutoFillBackground(False)
        self.pushButton3.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton3.setAutoRepeat(False)
        self.pushButton3.setObjectName("pushButton3")
        
        quit = QAction("Quit", self)
        quit.triggered.connect(self.closeEvent)


        _translate = QtCore.QCoreApplication.translate
        self.stack1.setWindowTitle(_translate("self.stack1", "Easy AFS"))
        self.pushButton1.setText(_translate("self.stack1", "Исходные данные"))
        self.pushButton2.setText(_translate("self.stack1", "Паспорта АФС"))
        self.pushButton3.setText(_translate("self.stack1", "Журнал полётов"))

        QtCore.QMetaObject.connectSlotsByName(self.stack1)




        
    

####################################################################################################
##########
#                                  страница с исходными данными
##########
#####################################################################################################
    def Window2UI(self):
        self.stack2.setObjectName("MainWindow")
        self.stack2.resize(1200, 800)
        self.stack2.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack2.sizePolicy().hasHeightForWidth())
        self.stack2.setSizePolicy(sizePolicy)
        self.stack2.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack2.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack2.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack2)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 51))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_4.setGeometry(QtCore.QRect(40, 360, 381, 91))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_5.setGeometry(QtCore.QRect(40, 430, 381, 91))
        self.textBrowser_5.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_folder = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_folder.setGeometry(QtCore.QRect(60, 150, 571, 31))
        self.plainTextEdit_folder.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_folder.setObjectName("plainTextEdit_folder")
        self.plainTextEdit_folder_2 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_folder_2.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_folder_2.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_folder_2.setObjectName("plainTextEdit_folder_2")
        self.plainTextEdit_folder_3 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_folder_3.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_folder_3.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_folder_3.setObjectName("plainTextEdit_folder_3")
        self.plainTextEdit_folder_4 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_folder_4.setGeometry(QtCore.QRect(60, 360, 571, 31))
        self.plainTextEdit_folder_4.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_folder_4.setObjectName("plainTextEdit_folder_4")
        self.plainTextEdit_folder_5 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_folder_5.setGeometry(QtCore.QRect(60, 430, 571, 31))
        self.plainTextEdit_folder_5.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_folder_5.setObjectName("plainTextEdit_folder_5")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_6.setGeometry(QtCore.QRect(250, 60, 181, 51))
        self.textBrowser_6.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_6.setObjectName("textBrowser_6")
        self.pushButton_download = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_download.setGeometry(QtCore.QRect(250, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.pushButton_download.setFont(font)
        self.pushButton_download.setMouseTracking(True)
        self.pushButton_download.setTabletTracking(False)
        self.pushButton_download.setFocusPolicy(QtCore.Qt.WheelFocus)
        self.pushButton_download.setAutoFillBackground(False)
        self.pushButton_download.setStyleSheet("QPushButton{background-color: rgb(0, 151, 216);\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"border-radius: 5px;}\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_download.setAutoRepeat(False)
        self.pushButton_download.setObjectName("pushButton")
        self.pushButton_folder = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_folder.setGeometry(QtCore.QRect(650, 150, 31, 31))
        self.pushButton_folder.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_folder.setObjectName("pushButton_folder")
        self.pushButton_folder_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_folder_2.setGeometry(QtCore.QRect(650, 360, 31, 31))
        self.pushButton_folder_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_folder_2.setObjectName("pushButton_folder_2")

        _translate = QtCore.QCoreApplication.translate
        self.stack2.setWindowTitle(_translate("self.stack2", "Easy AFS"))
        self.textBrowser.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Выберите папку с логами</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Введите ФИО</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Введите наименование объекта</span></p></body></html>"))
        self.textBrowser_4.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Выберите путь к документу excel</span></p></body></html>"))
        self.textBrowser_5.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Имя документа excel</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("self.stack2", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Исходные данные</span></p></body></html>"))
        self.pushButton_download.setText(_translate("self.stack2", "Загрузить"))
        self.pushButton_folder.setText(_translate("self.stack2", "..."))
        self.pushButton_folder_2.setText(_translate("self.stack2", "..."))

        QtCore.QMetaObject.connectSlotsByName(self.stack2)



####################################################################################################
##########
#                                  страница с паспортом АФС 1
##########
#####################################################################################################


    def Window3UI(self):
        self.stack3.setObjectName("Pasport_AFS_1")
        self.stack3.resize(1200, 800)
        self.stack3.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack3.sizePolicy().hasHeightForWidth())
        self.stack3.setSizePolicy(sizePolicy)
        self.stack3.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack3.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack3.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack3)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 51))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_5.setGeometry(QtCore.QRect(40, 430, 381, 51))
        self.textBrowser_5.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_6.setGeometry(QtCore.QRect(40, 360, 331, 51))
        self.textBrowser_6.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_6.setObjectName("textBrowser_6")
        self.textBrowser_7 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_7.setGeometry(QtCore.QRect(40, 500, 231, 51))
        self.textBrowser_7.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_7.setObjectName("textBrowser_7")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_data = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_data.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_data.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_data.setObjectName("plainTextEdit_data")
        self.plainTextEdit_time = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_time.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_time.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_time.setObjectName("plainTextEdit_time")
        self.pushButton = QtWidgets.QPushButton(self.frame_2)
        self.pushButton.setGeometry(QtCore.QRect(120, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton.setFont(font)
        self.pushButton.setMouseTracking(True)
        self.pushButton.setAutoFillBackground(False)
        self.pushButton.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")



        self.pushButton.setAutoRepeat(False)
        self.pushButton.setObjectName("pushButton")
        self.comboBox = QtWidgets.QComboBox(self.frame_2)
        self.comboBox.setGeometry(QtCore.QRect(60, 150, 571, 31))
        self.comboBox.setAutoFillBackground(False)
        self.comboBox.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox.setEditable(False)
        self.comboBox.setCurrentText("")
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItems(["1","2","3","4","5","6","7","8","9","10","11","12","13","14","15"])
        self.comboBox_2 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_2.setGeometry(QtCore.QRect(60, 360, 571, 31))
        self.comboBox_2.setAutoFillBackground(False)
        self.comboBox_2.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_2.setEditable(False)
        self.comboBox_2.setCurrentText("")
        self.comboBox_2.setObjectName("comboBox_2")
        self.comboBox_2.addItems(["Площадная","Линейная","Объектная","Фасадная"])
        self.comboBox_3 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_3.setGeometry(QtCore.QRect(60, 430, 571, 31))
        self.comboBox_3.setAutoFillBackground(False)
        self.comboBox_3.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_3.setEditable(False)
        self.comboBox_3.setCurrentText("")
        self.comboBox_3.setObjectName("comboBox_3")
        self.comboBox_3.addItems(["Надир (2D)","Наклон (3D)"])
        self.comboBox_4 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_4.setGeometry(QtCore.QRect(60, 500, 571, 31))
        self.comboBox_4.setAutoFillBackground(False)
        self.comboBox_4.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_4.setEditable(False)
        self.comboBox_4.setCurrentText("")
        self.comboBox_4.setObjectName("comboBox_4")
        self.comboBox_4.addItems(["DJI Phantom 4 Pro","DJI Phantom 4 RTK","DJI Matrice 300 RTK","DJI Mavic 2 Pro","DJI Mavic 2 Enterprice Dual", "DJI Matrice 600 Pro", "Геоскан 401 геодезия", "Геоскан 101 геодезия"])
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_4.setGeometry(QtCore.QRect(250, 60, 241, 51))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.pushButton_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_2.setGeometry(QtCore.QRect(420, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setMouseTracking(True)
        self.pushButton_2.setAutoFillBackground(False)
        self.pushButton_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_2.setAutoRepeat(False)
        self.pushButton_2.setObjectName("pushButton_2")

        _translate = QtCore.QCoreApplication.translate
        self.stack3.setWindowTitle(_translate("MainWindow", "Easy AFS"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Номер полетного задания</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Дата полета</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Время полета</span></p></body></html>"))
        self.textBrowser_5.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Вид АФС</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Тип АФС</span></p></body></html>"))
        self.textBrowser_7.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Название БВС</span></p></body></html>"))
        self.pushButton.setText(_translate("MainWindow", "Назад"))
        self.textBrowser_4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Паспорт АФС</span></p></body></html>"))
        self.pushButton_2.setText(_translate("MainWindow", "Далее"))
        QtCore.QMetaObject.connectSlotsByName(self.stack3)




####################################################################################################
##########
#                                  страница с паспортом АФС 2
##########
#####################################################################################################



    def Window4UI(self):
        self.stack4.setObjectName("MainWindow")
        self.stack4.resize(1200, 800)
        self.stack4.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack4.sizePolicy().hasHeightForWidth())
        self.stack4.setSizePolicy(sizePolicy)
        self.stack4.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack4.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack4.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack4)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 51))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_5.setGeometry(QtCore.QRect(40, 430, 381, 51))
        self.textBrowser_5.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_6.setGeometry(QtCore.QRect(40, 360, 331, 51))
        self.textBrowser_6.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_6.setObjectName("textBrowser_6")
        self.textBrowser_7 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_7.setGeometry(QtCore.QRect(40, 500, 231, 51))
        self.textBrowser_7.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_7.setObjectName("textBrowser_7")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_AFS_2 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_2.setGeometry(QtCore.QRect(60, 150, 571, 31))
        self.plainTextEdit_AFS_2.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_2.setObjectName("plainTextEdit_AFS_2")
        self.plainTextEdit_AFS_3 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_3.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_AFS_3.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_3.setObjectName("plainTextEdit_AFS_3")
        self.pushButton_AFS_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_AFS_2.setGeometry(QtCore.QRect(120, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_AFS_2.setFont(font)
        self.pushButton_AFS_2.setMouseTracking(True)
        self.pushButton_AFS_2.setAutoFillBackground(False)
        self.pushButton_AFS_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_AFS_2.setAutoRepeat(False)
        self.pushButton_AFS_2.setObjectName("pushButton")
        self.comboBox_AFS_2_1 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_AFS_2_1.setGeometry(QtCore.QRect(60, 430, 571, 31))
        self.comboBox_AFS_2_1.setAutoFillBackground(False)
        self.comboBox_AFS_2_1.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_AFS_2_1.setEditable(False)
        self.comboBox_AFS_2_1.setCurrentText("")
        self.comboBox_AFS_2_1.setObjectName("comboBox_AFS_2_1")
        self.comboBox_AFS_2_1.addItems(["PPK","RTK","Опознаки","RTK + Опознаки"])
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_4.setGeometry(QtCore.QRect(250, 60, 241, 51))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.plainTextEdit_AFS_4 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_4.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_AFS_4.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_4.setObjectName("plainTextEdit_AFS_4")
        self.plainTextEdit_AFS_5 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_5.setGeometry(QtCore.QRect(60, 360, 571, 31))
        self.plainTextEdit_AFS_5.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_5.setObjectName("plainTextEdit_AFS_5")
        self.plainTextEdit_AFS_6 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_6.setGeometry(QtCore.QRect(60, 500, 571, 31))
        self.plainTextEdit_AFS_6.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_6.setObjectName("plainTextEdit_AFS_6")
        self.pushButton_AFS_2_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_AFS_2_2.setGeometry(QtCore.QRect(420, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_AFS_2_2.setFont(font)
        self.pushButton_AFS_2_2.setMouseTracking(True)
        self.pushButton_AFS_2_2.setAutoFillBackground(False)
        self.pushButton_AFS_2_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_AFS_2_2.setAutoRepeat(False)
        self.pushButton_AFS_2_2.setObjectName("pushButton_2")
        #self.stack4.setCentralWidget(self.centralwidget)

        _translate = QtCore.QCoreApplication.translate
        self.stack4.setWindowTitle(_translate("MainWindow", "Easy AFS"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Регистрационный номер борта</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Полезная нагрузка 1</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Полезная нагрузка 2</span></p></body></html>"))
        self.textBrowser_5.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Метод решения</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">ПО для планирования полета</span></p></body></html>"))
        self.textBrowser_7.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Высота полета</span></p></body></html>"))
        self.pushButton_AFS_2.setText(_translate("MainWindow", "Назад"))
        self.textBrowser_4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Паспорт АФС</span></p></body></html>"))
        self.pushButton_AFS_2_2.setText(_translate("MainWindow", "Далее"))

        QtCore.QMetaObject.connectSlotsByName(self.stack4)


####################################################################################################
##########
#                                  страница с паспортом АФС 3
##########
#####################################################################################################


    def Window5UI(self):
        self.stack5.setObjectName("MainWindow")
        self.stack5.resize(1200, 800)
        self.stack5.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack5.sizePolicy().hasHeightForWidth())
        self.stack5.setSizePolicy(sizePolicy)
        self.stack5.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack5.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack5.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack5)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 51))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_5.setGeometry(QtCore.QRect(40, 430, 381, 51))
        self.textBrowser_5.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_6.setGeometry(QtCore.QRect(40, 360, 331, 51))
        self.textBrowser_6.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_6.setObjectName("textBrowser_6")
        self.textBrowser_7 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_7.setGeometry(QtCore.QRect(40, 500, 231, 51))
        self.textBrowser_7.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_7.setObjectName("textBrowser_7")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_AFS_3_1 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_3_1.setGeometry(QtCore.QRect(60, 150, 571, 31))
        self.plainTextEdit_AFS_3_1.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_3_1.setObjectName("plainTextEdit_folder_2")
        self.plainTextEdit_AFS_3_2 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_3_2.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_AFS_3_2.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_3_2.setObjectName("plainTextEdit_AFS_3_2")
        self.pushButton_AFS_3_1 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_AFS_3_1.setGeometry(QtCore.QRect(120, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_AFS_3_1.setFont(font)
        self.pushButton_AFS_3_1.setMouseTracking(True)
        self.pushButton_AFS_3_1.setAutoFillBackground(False)
        self.pushButton_AFS_3_1.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_AFS_3_1.setAutoRepeat(False)
        self.pushButton_AFS_3_1.setObjectName("pushButton")
        self.comboBox_AFS_3_1 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_AFS_3_1.setGeometry(QtCore.QRect(60, 430, 571, 31))
        self.comboBox_AFS_3_1.setAutoFillBackground(False)
        self.comboBox_AFS_3_1.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_AFS_3_1.setEditable(False)
        self.comboBox_AFS_3_1.setCurrentText("")
        self.comboBox_AFS_3_1.setObjectName("comboBox_AFS_3_1")
        self.comboBox_AFS_3_1.addItems(["Снег","Град","Дождь","Ливень","Морось","Туман"])
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_4.setGeometry(QtCore.QRect(250, 60, 241, 51))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.plainTextEdit_AFS_3_3 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_3_3.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_AFS_3_3.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_3_3.setObjectName("plainTextEdit_AFS_3_3")
        self.plainTextEdit_AFS_3_4 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_AFS_3_4.setGeometry(QtCore.QRect(60, 360, 571, 31))
        self.plainTextEdit_AFS_3_4.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_AFS_3_4.setObjectName("plainTextEdit_AFS_3_4")
        self.comboBox_AFS_3_2 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_AFS_3_2.setGeometry(QtCore.QRect(60, 500, 571, 31))
        self.comboBox_AFS_3_2.setAutoFillBackground(False)
        self.comboBox_AFS_3_2.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_AFS_3_2.setEditable(False)
        self.comboBox_AFS_3_2.setCurrentText("")
        self.comboBox_AFS_3_2.setObjectName("comboBox_AFS_3_2")
        self.comboBox_AFS_3_2.addItems(["Безоблачность","Переменная облачность","Облачность"])
        self.pushButton_AFS_3_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_AFS_3_2.setGeometry(QtCore.QRect(420, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_AFS_3_2.setFont(font)
        self.pushButton_AFS_3_2.setMouseTracking(True)
        self.pushButton_AFS_3_2.setAutoFillBackground(False)
        self.pushButton_AFS_3_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_AFS_3_2.setAutoRepeat(False)
        self.pushButton_AFS_3_2.setObjectName("pushButton_AFS_3_2")

        _translate = QtCore.QCoreApplication.translate
        self.stack5.setWindowTitle(_translate("self.stack5", "Easy AFS"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Продольное перекрытие</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Поперечное перекрытие</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Разрешение</span></p></body></html>"))
        self.textBrowser_5.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Осадки</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Количество снимков</span></p></body></html>"))
        self.textBrowser_7.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Облачность</span></p></body></html>"))
        self.pushButton_AFS_3_1.setText(_translate("MainWindow", "Назад"))
        self.textBrowser_4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Паспорт АФС</span></p></body></html>"))
        self.pushButton_AFS_3_2.setText(_translate("MainWindow", "Далее"))


        QtCore.QMetaObject.connectSlotsByName(self.stack5)

    

####################################################################################################
##########
#                                  страница с геодезией
##########
#####################################################################################################


    def Window6UI(self):
        self.stack6.setObjectName("self.stack6")
        self.stack6.resize(1200, 800)
        self.stack6.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack6.sizePolicy().hasHeightForWidth())
        self.stack6.setSizePolicy(sizePolicy)
        self.stack6.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack6.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack6.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack6)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 51))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.textBrowser_5 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_5.setGeometry(QtCore.QRect(40, 430, 381, 51))
        self.textBrowser_5.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_5.setObjectName("textBrowser_5")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.textBrowser_6 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_6.setGeometry(QtCore.QRect(40, 360, 331, 51))
        self.textBrowser_6.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_6.setObjectName("textBrowser_6")

        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_geodeziy_1 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_geodeziy_1.setGeometry(QtCore.QRect(60, 150, 571, 31))
        self.plainTextEdit_geodeziy_1.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_geodeziy_1.setObjectName("plainTextEdit_geodeziy_1")
        self.plainTextEdit_geodeziy_2 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_geodeziy_2.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_geodeziy_2.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_geodeziy_2.setObjectName("plainTextEdit_geodeziy_2")
        self.pushButton_geodeziy_1 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_geodeziy_1.setGeometry(QtCore.QRect(120, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_geodeziy_1.setFont(font)
        self.pushButton_geodeziy_1.setMouseTracking(True)
        self.pushButton_geodeziy_1.setAutoFillBackground(False)
        self.pushButton_geodeziy_1.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_geodeziy_1.setAutoRepeat(False)
        self.pushButton_geodeziy_1.setObjectName("pushButton")
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_4.setGeometry(QtCore.QRect(250, 60, 241, 51))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.plainTextEdit_geodeziy_3 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_geodeziy_3.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_geodeziy_3.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_geodeziy_3.setObjectName("plainTextEdit_geodeziy_3")
        self.plainTextEdit_geodeziy_4 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_geodeziy_4.setGeometry(QtCore.QRect(60, 360, 571, 31))
        self.plainTextEdit_geodeziy_4.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_geodeziy_4.setObjectName("plainTextEdit_geodeziy_4")
        self.plainTextEdit_geodeziy_5 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_geodeziy_5.setGeometry(QtCore.QRect(60, 430, 571, 31))
        self.plainTextEdit_geodeziy_5.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_geodeziy_5.setObjectName("plainTextEdit_geodeziy_5")

        self.pushButton_geodeziy_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_geodeziy_2.setGeometry(QtCore.QRect(420, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_geodeziy_2.setFont(font)
        self.pushButton_geodeziy_2.setMouseTracking(True)
        self.pushButton_geodeziy_2.setAutoFillBackground(False)
        self.pushButton_geodeziy_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_geodeziy_2.setAutoRepeat(False)
        self.pushButton_geodeziy_2.setObjectName("pushButton_geodeziy_2")
        
        _translate = QtCore.QCoreApplication.translate
        self.stack6.setWindowTitle(_translate("MainWindow", "Easy AFS"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Наименование точки(базы)</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Прибор(название, номер)</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Порядковый номер лога(базы)</span></p></body></html>"))
        self.textBrowser_5.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Название файла</span></p></body></html>"))
        self.textBrowser_6.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Высота прибора</span></p></body></html>"))
        self.pushButton_geodeziy_1.setText(_translate("MainWindow", "Назад"))
        self.textBrowser_4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Геодезия</span></p></body></html>"))
        self.pushButton_geodeziy_2.setText(_translate("MainWindow", "Далее"))

        QtCore.QMetaObject.connectSlotsByName(self.stack6)




####################################################################################################
##########
#                                  страница с паспортом АФС 2
##########
#####################################################################################################



    def Window7UI(self):
        self.stack7.setObjectName("self.stack7")
        self.stack7.resize(1200, 800)
        self.stack7.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack7.sizePolicy().hasHeightForWidth())
        self.stack7.setSizePolicy(sizePolicy)
        self.stack7.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack7.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack7.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack7)
        self.centralwidget.setObjectName("centralwidget")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 400, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(400, 800))
        self.frame.setMaximumSize(QtCore.QSize(400, 800))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.textBrowser = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser.setGeometry(QtCore.QRect(40, 150, 341, 51))
        self.textBrowser.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_2.setGeometry(QtCore.QRect(40, 220, 341, 61))
        self.textBrowser_2.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.frame)
        self.textBrowser_3.setGeometry(QtCore.QRect(40, 290, 361, 81))
        self.textBrowser_3.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(50, 40, 251, 71))
        self.label.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/ЦР-русский_-полный-_web_.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.frame)
        self.line.setGeometry(QtCore.QRect(390, 0, 16, 800))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.line.sizePolicy().hasHeightForWidth())
        self.line.setSizePolicy(sizePolicy)
        self.line.setMinimumSize(QtCore.QSize(0, 800))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(400, 0, 800, 800))
        self.frame_2.setMinimumSize(QtCore.QSize(800, 800))
        self.frame_2.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"}")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.plainTextEdit_primechania_1 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_primechania_1.setGeometry(QtCore.QRect(60, 290, 571, 31))
        self.plainTextEdit_primechania_1.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_primechania_1.setObjectName("plainTextEdit_primechania_1")
        self.pushButton_primechania_1 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_primechania_1.setGeometry(QtCore.QRect(120, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_primechania_1.setFont(font)
        self.pushButton_primechania_1.setMouseTracking(True)
        self.pushButton_primechania_1.setAutoFillBackground(False)
        self.pushButton_primechania_1.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_primechania_1.setAutoRepeat(False)
        self.pushButton_primechania_1.setObjectName("pushButton")
        self.comboBox_primechania = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_primechania.setGeometry(QtCore.QRect(60, 160, 571, 31))
        self.comboBox_primechania.setAutoFillBackground(False)
        self.comboBox_primechania.setStyleSheet("QComboBox{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}\n"
"\n"
"QComboBox::down-arrow{\n"
"    width: 10px;\n"
"    height: 10px;\n"
"    image : url(images/down.png);\n"
"}"
"QComboBox::drop-down {\n"
"    width: 20px;\n"
"    border-left-width: 1px;\n"
"    border-left-color: darkgray;\n"
"    border-left-style: solid; \n"
"    border-top-right-radius: 3px;\n"
"    border-bottom-right-radius: 3px;\n"
"}"
)
        self.comboBox_primechania.setEditable(False)
        self.comboBox_primechania.setCurrentText("")
        self.comboBox_primechania.setObjectName("comboBox_primechania")
        self.comboBox_primechania.addItems(["Да","Нет","Скорее да","Скорее нет"])
        self.textBrowser_4 = QtWidgets.QTextBrowser(self.frame_2)
        self.textBrowser_4.setGeometry(QtCore.QRect(250, 60, 241, 51))
        self.textBrowser_4.setStyleSheet("QTextBrowser{\n"
"    \n"
"    border: 0px\n"
"}")
        self.textBrowser_4.setObjectName("textBrowser_4")
        self.plainTextEdit_primechania_2 = QtWidgets.QPlainTextEdit(self.frame_2)
        self.plainTextEdit_primechania_2.setGeometry(QtCore.QRect(60, 220, 571, 31))
        self.plainTextEdit_primechania_2.setStyleSheet("QPlainTextEdit{\n"
"    \n"
"    background-color: rgb(255, 255, 255);\n"
"    border-radius: 10px;\n"
"    border: 1px solid black;\n"
"}")
        self.plainTextEdit_primechania_2.setObjectName("plainTextEdit_primechania_2")
        self.pushButton_primechania_2 = QtWidgets.QPushButton(self.frame_2)
        self.pushButton_primechania_2.setGeometry(QtCore.QRect(420, 620, 211, 51))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_primechania_2.setFont(font)
        self.pushButton_primechania_2.setMouseTracking(True)
        self.pushButton_primechania_2.setAutoFillBackground(False)
        self.pushButton_primechania_2.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_primechania_2.setAutoRepeat(False)
        self.pushButton_primechania_2.setObjectName("pushButton_primechania_2")
 

        _translate = QtCore.QCoreApplication.translate
        self.stack7.setWindowTitle(_translate("MainWindow", "Easy AFS"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Использование полета в обработке</span></p></body></html>"))
        self.textBrowser_2.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Причина, по которой нельзя использовать</span></p></body></html>"))
        self.textBrowser_3.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Происшествия</span></p></body></html>"))
        self.pushButton_primechania_1.setText(_translate("MainWindow", "Назад"))
        self.textBrowser_4.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:12pt;\">Примечания</span></p></body></html>"))
        self.pushButton_primechania_2.setText(_translate("MainWindow", "Далее"))
        
        QtCore.QMetaObject.connectSlotsByName(self.stack7)


####################################################################################################
##########
#                                  страница с журналом АФС
##########
#####################################################################################################



    def Window8UI(self):
        self.stack8.setObjectName("self.stack8")
        self.stack8.resize(1200, 800)
        self.stack8.setWindowIcon(QIcon("images/logo.png"))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.stack8.sizePolicy().hasHeightForWidth())
        self.stack8.setSizePolicy(sizePolicy)
        self.stack8.setMinimumSize(QtCore.QSize(1200, 800))
        self.stack8.setMaximumSize(QtCore.QSize(1200, 800))
        self.stack8.setSizeIncrement(QtCore.QSize(1200, 800))
        self.centralwidget = QtWidgets.QWidget(self.stack8)
        self.centralwidget.setObjectName("centralwidget")
        
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0, 1203, 90))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.frame.sizePolicy().hasHeightForWidth())
        self.frame.setSizePolicy(sizePolicy)
        self.frame.setMinimumSize(QtCore.QSize(1203, 90))
        self.frame.setMaximumSize(QtCore.QSize(1203, 90))
        self.frame.setStyleSheet("QFrame{\n"
"    \n"
"    background-color: rgb(228, 232, 235);\n"
"}")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(9)
        
        self.label1 = QtWidgets.QLabel(self.frame)
        self.label1.setGeometry(QtCore.QRect(0, 0, 127, 87))
        self.label1.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label1.setStyleSheet("padding :5px")
        self.label1.setText("АФС")
        self.label1.setScaledContents(True)
        self.label1.setObjectName("label1")
        self.label1.setFont(font)
        self.label1.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label2 = QtWidgets.QLabel(self.frame)
        self.label2.setGeometry(QtCore.QRect(127, 0, 160, 87))
        self.label2.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label2.setStyleSheet("padding :5px")
        self.label2.setText("Дата полёта")
        self.label2.setScaledContents(True)
        self.label2.setObjectName("label2")
        self.label2.setFont(font)
        self.label2.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label3 = QtWidgets.QLabel(self.frame)
        self.label3.setGeometry(QtCore.QRect(287, 0, 122, 87))
        self.label3.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label3.setStyleSheet("padding :5px")
        self.label3.setText("БВС")
        self.label3.setScaledContents(True)
        self.label3.setObjectName("label3")
        self.label3.setFont(font)
        self.label3.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label4 = QtWidgets.QLabel(self.frame)
        self.label4.setGeometry(QtCore.QRect(409, 0, 100, 87))
        self.label4.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label4.setStyleSheet("padding :5px")
        self.label4.setText("№ борта")
        self.label4.setScaledContents(True)
        self.label4.setObjectName("label4")
        self.label4.setFont(font)
        self.label4.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label5 = QtWidgets.QLabel(self.frame)
        self.label5.setGeometry(QtCore.QRect(509, 0, 110, 87))
        self.label5.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label5.setStyleSheet("padding :5px")
        self.label5.setText("№ полётного задания")
        self.label5.setWordWrap(True)
        self.label5.setScaledContents(True)
        self.label5.setObjectName("label5")
        self.label5.setFont(font)
        self.label5.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label6 = QtWidgets.QLabel(self.frame)
        self.label6.setGeometry(QtCore.QRect(619, 0, 209, 87))
        self.label6.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label6.setStyleSheet("padding :5px")
        self.label6.setText("Наименование объекта")
        self.label6.setWordWrap(True)
        self.label6.setScaledContents(True)
        self.label6.setObjectName("label6")
        self.label6.setFont(font)
        self.label6.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label7 = QtWidgets.QLabel(self.frame)
        self.label7.setGeometry(QtCore.QRect(828, 0, 150, 87))
        self.label7.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label7.setStyleSheet("padding :5px")
        self.label7.setText("Тип съемки")
        self.label7.setWordWrap(True)
        self.label7.setScaledContents(True)
        self.label7.setObjectName("label7")
        self.label7.setFont(font)
        self.label7.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label8 = QtWidgets.QLabel(self.frame)
        self.label8.setGeometry(QtCore.QRect(978, 0, 100, 87))
        self.label8.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label8.setStyleSheet("padding :5px")
        self.label8.setText("Вид съемки")
        self.label8.setWordWrap(True)
        self.label8.setScaledContents(True)
        self.label8.setObjectName("label8")
        self.label8.setFont(font)
        self.label8.setAlignment(QtCore.Qt.AlignCenter)
        
        self.label9 = QtWidgets.QLabel(self.frame)
        self.label9.setGeometry(QtCore.QRect(1078, 0, 122, 87))
        self.label9.setFrameShadow(QtWidgets.QFrame.Plain)
        self.label9.setStyleSheet("padding :5px")
        self.label9.setText("Подробнее")
        self.label9.setWordWrap(True)
        self.label9.setScaledContents(True)
        self.label9.setObjectName("label9")
        self.label9.setFont(font)
        self.label9.setAlignment(QtCore.Qt.AlignCenter)

        self.pushButton_jurnal = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_jurnal.setGeometry(QtCore.QRect(425, 620, 389, 71))
        font = QtGui.QFont()
        font.setFamily("Rockwell")
        font.setPointSize(10)
        self.pushButton_jurnal.setFont(font)
        self.pushButton_jurnal.setMouseTracking(True)
        self.pushButton_jurnal.setAutoFillBackground(False)
        self.pushButton_jurnal.setStyleSheet("QPushButton{\n"
"    background-color: rgb(0, 151, 216);\n"
"    border-radius: 10px;\n"
"    border-color: rgb(0, 0, 0);\n"
"border: 1px solid;\n"
"}\n"
"\n"
"\n"
"\n"
"QPushButton:pressed\n"
"{\n"
"background-color: rgb(0, 111, 186);\n"
"border-style:solid;\n"
"border-width:2px;\n"
"}@")
        self.pushButton_jurnal.setAutoRepeat(False)
        self.pushButton_jurnal.setObjectName("pushButton1")
        

        font_1 = QtGui.QFont()
        font_1.setFamily("Rockwell")
        font_1.setPointSize(9)
        
        # афс должны идти в порядке возрастания

        def foo():
                with open("data.json", "r") as read_file:
                        data = json.load(read_file)
                x=100
                y=180
                for i in data:
                        if i=='pasport_ishodnie_dannye':
                                pass
                        else:    
                                self.frame_1 = QtWidgets.QFrame(self.centralwidget)
                                self.frame_1.setGeometry(QtCore.QRect(0, x, 1203, y))
                                sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
                                sizePolicy.setHorizontalStretch(0)
                                sizePolicy.setVerticalStretch(0)
                                sizePolicy.setHeightForWidth(self.frame_1.sizePolicy().hasHeightForWidth())
                                self.frame_1.setSizePolicy(sizePolicy)
                                self.frame_1.setMinimumSize(QtCore.QSize(1203, 90))
                                self.frame_1.setMaximumSize(QtCore.QSize(1203, 90))
                                self.frame_1.setStyleSheet("QFrame{\n"
                        "    \n"
                        "    background-color: rgb(228, 232, 235);\n"
                        "}")
                                self.frame_1.setFrameShape(QtWidgets.QFrame.StyledPanel)
                                self.frame_1.setFrameShadow(QtWidgets.QFrame.Raised)
                                self.frame_1.setObjectName("frame_1")

                                self.label_1 = QtWidgets.QLabel(self.frame_1)
                                self.label_1.setGeometry(QtCore.QRect(0, 0, 127, 87))
                                self.label_1.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_1.setStyleSheet("padding :5px")
                                self.label_1.setText("АФС"+ data[i]["Mission_number"])
                                self.label_1.setScaledContents(True)
                                self.label_1.setObjectName("label_1")
                                self.label_1.setFont(font_1)
                                self.label_1.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_2 = QtWidgets.QLabel(self.frame_1)
                                self.label_2.setGeometry(QtCore.QRect(127, 0, 150, 87))
                                self.label_2.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_2.setStyleSheet("padding :5px")
                                self.label_2.setText(data[i]["Date"])
                                self.label_2.setScaledContents(True)
                                self.label_2.setObjectName("label2")
                                self.label_2.setFont(font_1)
                                self.label_2.setAlignment(QtCore.Qt.AlignCenter)
                                
                                self.label_3 = QtWidgets.QLabel(self.frame_1)
                                self.label_3.setGeometry(QtCore.QRect(220, 0, 287, 87))
                                self.label_3.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_3.setStyleSheet("padding :5px")
                                self.label_3.setText(data[i]["UMA_name"])
                                self.label_3.setScaledContents(True)
                                self.label_3.setObjectName("label3")
                                self.label_3.setFont(font_1)
                                self.label_3.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_4 = QtWidgets.QLabel(self.frame_1)
                                self.label_4.setGeometry(QtCore.QRect(409, 0, 100, 87))
                                self.label_4.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_4.setStyleSheet("padding :5px")
                                self.label_4.setText(data[i]["registry_number"])
                                self.label_4.setScaledContents(True)
                                self.label_4.setObjectName("label_4")
                                self.label_4.setFont(font_1)
                                self.label_4.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_5 = QtWidgets.QLabel(self.frame_1)
                                self.label_5.setGeometry(QtCore.QRect(509, 0, 110, 87))
                                self.label_5.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_5.setStyleSheet("padding :5px")
                                self.label_5.setText(data[i]["Mission_number"])
                                self.label_5.setWordWrap(True)
                                self.label_5.setScaledContents(True)
                                self.label_5.setObjectName("label5")
                                self.label_5.setFont(font_1)
                                self.label_5.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_6 = QtWidgets.QLabel(self.frame_1)
                                self.label_6.setGeometry(QtCore.QRect(619, 0, 209, 87))
                                self.label_6.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_6.setStyleSheet("padding :5px")
                                self.label_6.setText(data["pasport_ishodnie_dannye"]["object_name"])
                                self.label_6.setWordWrap(True)
                                self.label_6.setScaledContents(True)
                                self.label_6.setObjectName("label6")
                                self.label_6.setFont(font_1)
                                self.label_6.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_7 = QtWidgets.QLabel(self.frame_1)
                                self.label_7.setGeometry(QtCore.QRect(828, 0, 150, 87))
                                self.label_7.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_7.setStyleSheet("padding :5px")
                                self.label_7.setText(data[i]["AFS_type"])
                                self.label_7.setWordWrap(True)
                                self.label_7.setScaledContents(True)
                                self.label_7.setObjectName("label7")
                                self.label_7.setFont(font_1)
                                self.label_7.setAlignment(QtCore.Qt.AlignCenter)

                                self.label_8 = QtWidgets.QLabel(self.frame_1)
                                self.label_8.setGeometry(QtCore.QRect(978, 0, 100, 87))
                                self.label_8.setFrameShadow(QtWidgets.QFrame.Plain)
                                self.label_8.setStyleSheet("padding :5px")
                                self.label_8.setText(data[i]["AFS_mode"])
                                self.label_8.setWordWrap(True)
                                self.label_8.setScaledContents(True)
                                self.label_8.setObjectName("label8")
                                self.label_8.setFont(font_1)
                                self.label_8.setAlignment(QtCore.Qt.AlignCenter)


                                # добавление к координатам для следующего значения
                                x+=100
                                y+=100
                

        
        self.pushButton3.clicked.connect(foo)
        

        
        
                        



        _translate = QtCore.QCoreApplication.translate
        self.stack8.setWindowTitle(_translate("MainWindow", "Easy AFS"))
        self.pushButton_jurnal.setText(_translate("MainWindow", "В меню"))
        QtCore.QMetaObject.connectSlotsByName(self.stack8)



