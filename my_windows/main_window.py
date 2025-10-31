# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'main_window.ui'
##
## Created by: Qt User Interface Compiler version 6.6.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,Qt,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QHeaderView, QLabel, QMainWindow, QMenu,
    QMenuBar, QPushButton, QSizePolicy, QStatusBar,QLineEdit, QComboBox,  QInputDialog,  
    QTableView, QTableWidget, QTableWidgetItem, QWidget, QVBoxLayout,)
import global_vars 

# from my_threads.show_result_table_sheet import ShowResultTableThread
# from my_threads.open_files_from_marked_folder import OpenFilesFromMarkedFolderThread


class Ui_MainWindow(object):


    def setupUi(self, MainWindow):
        #MainWindow.setFixedWidth(1366) 
        #MainWindow.setFixedHeight(768) 
        MainWindow.resize(880, 320)
        # MainWindow.setMaximumSize(720, 360)
        MainWindow.setMinimumSize(350, 340)    
        MainWindow.setWindowTitle(f"Добавляем заказ по номерам отправки, контейнера, вагона (ver.{global_vars.version})")
       

        self.centralwidget = QWidget(MainWindow)

###########################################################################################################################
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setGeometry(QRect(0, 0, 720, 28))

        self.menu = QMenu(self.menubar)
        self.menu.setTitle("Подключение к DWH")
        self.menubar.addAction(self.menu.menuAction())           

        self.action_log_in = QAction(MainWindow)
        self.action_log_in.setText("Подключиться")  
        self.menu.addAction(self.action_log_in)

        self.action_log_in_check = QAction(MainWindow)
        self.action_log_in_check.setText("Проверить подключение")
        self.menu.addAction(self.action_log_in_check)         

        self.action_log_out = QAction(MainWindow)
        self.action_log_out.setText("Отключиться")
        self.menu.addAction(self.action_log_out)  
     
        self.action_show_manual = QAction(self.menubar)
        self.action_show_manual.setText("Инструкция") 
        self.menubar.addAction(self.action_show_manual)  

        self.action_show_dev_info = QAction(self.menubar)
        self.action_show_dev_info.setText("Связь с разработчиками") 
        self.menubar.addAction(self.action_show_dev_info) 

###########################################################################################################################
        self.verticalLayoutWidgetButtons = QWidget(self.centralwidget)
        self.verticalLayoutWidgetButtons.setGeometry(QRect(10, 42, 320, 210))
        self.verticalLayoutButtons = QVBoxLayout(self.verticalLayoutWidgetButtons)
        self.verticalLayoutButtons.setContentsMargins(10, 0, 0, 0)    
             
        self.pushButtonChooseProjectFolder = QPushButton("Выберите папку проекта")
        self.pushButtonChooseProjectFolder.setEnabled(True)
        self.verticalLayoutButtons.addWidget(self.pushButtonChooseProjectFolder)
        
        self.pushButtonXLStoXLSX = QPushButton("Конвертировать xls и xlsm в xlsx")
        self.pushButtonXLStoXLSX.setEnabled(False)
        self.verticalLayoutButtons.addWidget(self.pushButtonXLStoXLSX)

        self.pushButtonProcessing = QPushButton("Обработка")
        self.pushButtonProcessing.setEnabled(False)
        self.verticalLayoutButtons.addWidget(self.pushButtonProcessing)

        self.pushButtonOpenChoosedFiles = QPushButton("Открыть выбранные файлы из папки .Исходники")
        self.pushButtonOpenChoosedFiles.setEnabled(False)        
        self.verticalLayoutButtons.addWidget(self.pushButtonOpenChoosedFiles)
 
        self.pushButtonOpenChoosedMDFiles = QPushButton("Открыть выбранные файлы из папки .Размеченные")
        self.pushButtonOpenChoosedMDFiles.setEnabled(False)        
        self.verticalLayoutButtons.addWidget(self.pushButtonOpenChoosedMDFiles)
 
        self.pushButtonDelChoosedMDFiles = QPushButton("Удалить выбранные файлы из папки .Размеченные")
        self.pushButtonDelChoosedMDFiles.setEnabled(False)
        self.pushButtonDelChoosedMDFiles.setStyleSheet("color: red")        
        self.verticalLayoutButtons.addWidget(self.pushButtonDelChoosedMDFiles)
 
        self.pushButtonConcat = QPushButton("Объединить")
        self.pushButtonConcat.setEnabled(False)        
        self.verticalLayoutButtons.addWidget(self.pushButtonConcat)        

        self.pushButtonMakeFiles = QPushButton("Создать файлы для 1-С")
        self.pushButtonMakeFiles.setEnabled(False)        
        self.verticalLayoutButtons.addWidget(self.pushButtonMakeFiles)    

        
        #####################################################################################
        self.verticalLayoutWidgetLabels = QWidget(self.centralwidget)
        self.verticalLayoutWidgetLabels.setGeometry(QRect(10, 254, 1366, 68)) #QRect(10, 120, 320, 120)
        self.verticalLayoutLabels = QVBoxLayout(self.verticalLayoutWidgetLabels)
        self.verticalLayoutLabels.setContentsMargins(10, 10, 10, 10)    


        self.project_folder_label = QLabel('Не выбрана папка проекта.')
        self.project_folder_label.setStyleSheet('color: red')      
        self.verticalLayoutLabels.addWidget(self.project_folder_label)

        self.info_label = QLabel('Выберите папу проекта!')
        self.info_label.setStyleSheet('color: red')        
        self.verticalLayoutLabels.addWidget(self.info_label)

        self.login_label = QLabel('Подключитесь к DWH!')
        self.login_label.setStyleSheet('color: red')
        self.verticalLayoutLabels.addWidget(self.login_label)
