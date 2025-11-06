from PySide6 import QtWidgets, QtCore
from colorama import Fore
import global_vars 
import os
import pyperclip



class OpenChoosedFilesThread(QtCore.QThread):
    def __init__ (self, md_files=False, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.message_title = "Открываем выбранные файлы:"
        self.md_files = md_files

    def run(self): 
        files_list_in_pyperclip = set(pyperclip.paste().splitlines())

        self.error_message = ""
        self.warning_message = ""
        self.info_message = ""

        
        folder = os.path.join(global_vars.project_folder,'.Размеченные') if self.md_files \
            else os.path.join(global_vars.project_folder,'.Исходники')

        for file in files_list_in_pyperclip:
            file = file if self.md_files else file[3:]
            file_to_start = os.path.join(folder, file)
            if not os.path.exists(file_to_start):
                self.error_message = f"Файл {file_to_start} не найден в папке .Размеченные/"
                return
            os.startfile(file_to_start)

    def on_clicked(self):      
        self.start() # Запускаем поток  
     
        
    def on_started(self): # Вызывается при запуске потока     
        pass


    def on_finished(self): # Вызывается при завершении потока
        if self.error_message:
      
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText(self.error_message.replace('\n',' '))

            QtWidgets.QMessageBox.critical(None,
                self.message_title,
                self.error_message,
                buttons=QtWidgets.QMessageBox.StandardButton.Ok)
            
        if self.warning_message:
            QtWidgets.QMessageBox.warning(None,
                self.message_title,
                self.warning_message,
                buttons=QtWidgets.QMessageBox.StandardButton.Ok)
            
        if self.info_message:
            QtWidgets.QMessageBox.information(None,
                self.message_title,
                self.info_message,
                buttons=QtWidgets.QMessageBox.StandardButton.Ok)
