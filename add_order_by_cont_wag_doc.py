from PySide6 import QtWidgets, QtCore
from my_windows import main_window, log_in_dialog
from my_functions.dwh import log_out, get_params, connection_settings_file_creator

import sys, os
import global_vars
from colorama import Fore

from my_threads.log_in_check import LogInCheck
from my_threads.functions import clean_process_folder
from my_threads.choose_project_folder import ChooseProjectFolderThread
from my_threads.xls_to_xlsx import XLS_TO_xlsxThread
from my_threads.processing import ProcessingThread
from my_threads.open_choosed_files import OpenChoosedFilesThread
from my_threads.del_choosed_md_files import DelChoosedMDFilesThread
from my_threads.concat import ConcatThread

class LogInDialog(QtWidgets.QWidget):
    def __init__(self, parent):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = log_in_dialog.Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.buttonBox.accepted.connect(self.accept)
        self.log_in_check_thread.start()  
        params = get_params()
        self.ui.lineEditHostField.setText(params[0])         
        self.ui.lineEditPortField.setText(params[1])         
        self.ui.lineEditDBNameField.setText(params[2])         
        self.ui.lineEditUserField.setText(params[3])         
        self.ui.lineEditPasswordField.setText(params[4])  
    def accept(self):
        connection_settings_file_creator(self.ui.lineEditHostField.text(),
               self.ui.lineEditPortField.text(),
               self.ui.lineEditDBNameField.text(),
               self.ui.lineEditUserField.text(),
               self.ui.lineEditPasswordField.text()
               )
        self.log_in_check_thread.start()

    #######################################################################
    #######################################################################   
    log_in_check_thread = LogInCheck() # создаём поток проверки подключения     



class MyWindow(QtWidgets.QWidget):
    def __init__ (self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)

        global_vars.ui = main_window.Ui_MainWindow()
        global_vars.ui.setupUi(self)
        self.log_in_check_thread.start()

        self.processing_thread.mysignal.connect(self.processing_thread.on_signal, QtCore.Qt.ConnectionType.QueuedConnection)
        self.concat_thread.mysignal.connect(self.processing_thread.on_signal, QtCore.Qt.ConnectionType.QueuedConnection)       

        global_vars.ui.action_log_in.triggered.connect(self.show_log_in_dialog)
        global_vars.ui.action_log_in_check.triggered.connect(lambda: self.log_in_check_thread.start())          
        global_vars.ui.action_log_out.triggered.connect(log_out)    

        global_vars.ui.action_show_manual.triggered.connect(self.show_manual)   
        global_vars.ui.action_show_dev_info.triggered.connect(self.show_dev_info) 
            
        global_vars.ui.pushButtonChooseProjectFolder.clicked.connect(self.choose_project_folder_thread.on_clicked)
        self.choose_project_folder_thread.started.connect(self.choose_project_folder_thread.on_started)
        self.choose_project_folder_thread.finished.connect(self.choose_project_folder_thread.on_finished)
        
        global_vars.ui.pushButtonXLStoXLSX.clicked.connect(self.xls_to_xlsx_thread.on_clicked)
        self.xls_to_xlsx_thread.started.connect(self.xls_to_xlsx_thread.on_started)
        self.xls_to_xlsx_thread.finished.connect(self.xls_to_xlsx_thread.on_finished)         
         
        global_vars.ui.pushButtonProcessing.clicked.connect(self.processing_thread.on_clicked)
        self.processing_thread.started.connect(self.processing_thread.on_started)
        self.processing_thread.finished.connect(self.processing_thread.on_finished)

        global_vars.ui.pushButtonOpenChoosedFiles.clicked.connect(self.open_choosed_files_thread.on_clicked)
        self.open_choosed_files_thread.started.connect(self.open_choosed_files_thread.on_started)
        self.open_choosed_files_thread.finished.connect(self.open_choosed_files_thread.on_finished)

        global_vars.ui.pushButtonOpenChoosedMDFiles.clicked.connect(self.open_choosed_mdfiles_thread.on_clicked)
        self.open_choosed_mdfiles_thread.started.connect(self.open_choosed_mdfiles_thread.on_started)
        self.open_choosed_mdfiles_thread.finished.connect(self.open_choosed_mdfiles_thread.on_finished)

        global_vars.ui.pushButtonDelChoosedMDFiles.clicked.connect(self.del_choosed_md_files_thread.on_clicked)
        self.del_choosed_md_files_thread.started.connect(self.del_choosed_md_files_thread.on_started)
        self.del_choosed_md_files_thread.finished.connect(self.del_choosed_md_files_thread.on_finished)

        global_vars.ui.pushButtonConcat.clicked.connect(self.concat_thread.on_clicked)
        self.concat_thread.started.connect(self.concat_thread.on_started)
        self.concat_thread.finished.connect(self.concat_thread.on_finished)        

    def show_log_in_dialog(self):
        self.login_dialog_window = LogInDialog(parent = None)
        # self.login_dialog_window.setWindowModality(True)
        self.login_dialog_window.show()

    def show_dev_info(self):
        QtWidgets.QMessageBox.about(None, "Контакты разработчиков", global_vars.dev_info)

    def show_manual(self):
        QtWidgets.QMessageBox.about(None, "Инструкция", global_vars.manual)    


     
####################################################################################
####################################################################################   

    choose_project_folder_thread = ChooseProjectFolderThread() 

####################################################################################
####################################################################################  

    xls_to_xlsx_thread = XLS_TO_xlsxThread() 

####################################################################################
####################################################################################  

    log_in_check_thread = LogInCheck()        # создаём поток проверки подключения 
    processing_thread = ProcessingThread() 
    open_choosed_files_thread = OpenChoosedFilesThread(md_files = False)  
    open_choosed_mdfiles_thread = OpenChoosedFilesThread(md_files = True)      
    del_choosed_md_files_thread = DelChoosedMDFilesThread()
    concat_thread = ConcatThread()      

####################################################################################
####################################################################################  

         

if __name__ == "__main__":


    # подчищаем папку .Обработка из папки проекта
    # выбранной при предыдущем запуске программы
    if os.path.exists('.session_folder'):
        with open('.session_folder', encoding='utf-8') as f:
            precending_project_folder = f.readline()
    else:
        precending_project_folder = ''
    clean_process_folder(precending_project_folder)

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MyWindow()
    window.show()

    sys.exit(app.exec())
  
    