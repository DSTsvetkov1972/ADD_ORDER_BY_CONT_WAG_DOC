from PySide6 import QtWidgets, QtCore
from colorama import Fore
import global_vars 
import os
import pandas as pd
from openpyxl import Workbook,load_workbook,styles
from copy import copy
from time import sleep
from tkinter import messagebox
from colorama import Fore

class ChooseProjectFolderThread(QtCore.QThread):
    def __init__ (self, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.message_title = "Выбор папки проекта:"

    def run(self): 
        self.error_message = ""
        self.warning_message = ""

        print(Fore.YELLOW, os.path.exists(os.path.join(global_vars.project_folder,'.Исходники')), Fore.RESET)

        if not global_vars.project_folder:

            global_vars.ui.project_folder_label.setStyleSheet('color: red')  
            global_vars.ui.project_folder_label.setText('Папка проекта: не выбрана') 
          
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText('Выберите папку проекта')
                
            self.error_message = ('Папка проекта: не выбрана')
            return
 
        if os.path.exists(os.path.join(global_vars.project_folder,'.Исходники')):
            source_files_list = list(os.walk(os.path.join(global_vars.project_folder,'.Исходники')))[0][2]
            source_excels_list = [file for file in source_files_list if file[-4:] == 'xlsx']            
            source_old_excels_list = [file for file in source_files_list if file[-4:] in ['.xls', 'xlsm']]

        else:


            global_vars.ui.project_folder_label.setStyleSheet('color: red')  
            global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 
          
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText(f'В папке проекта нет папки .Исходники/.')
                
            self.error_message = ('В папке проекта нет папки .Исходники!\n'
                                  'Создайте в папке проекта папку .Исходники\n'
                                  'и скопируйте в неё файлы, которые нужно обработать,\n'
                                  'затем снова нажмите кнопку "Выбирете папку проекта"!')

            global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)
            global_vars.ui.pushButtonProcessing.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)              
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False)  
            global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(False)             
            global_vars.ui.pushButtonConcat.setEnabled(False)             
            return 
        
        if not source_files_list:

            global_vars.ui.project_folder_label.setStyleSheet('color: red')  
            global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 
       
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText(f'В папке проекта есть папка .Исходники/, но она не содержит файлов.')
                             
            self.warning_message = f'Папка .Исходники/ не содержит файлов!\nСкопируйте в папку .Исходники/ файлы для обработки и снова нажмите кнопку "Выберите папку проекта"!'

            global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)
            global_vars.ui.pushButtonProcessing.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)            
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False)
            global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(False)                 
            global_vars.ui.pushButtonConcat.setEnabled(False)       
            return                  

        if source_old_excels_list:       

            global_vars.ui.project_folder_label.setStyleSheet('color: red')  
            global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 
       
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText('В папке проекта есть папка .Исходники/, но в ней некоторые файлы в формате .xls или .xlsm')
                             
            self.warning_message = 'В папке проекта есть папка .Исходники/, но в ней некоторые файлы в формате .xls или .xlsm'

            global_vars.ui.pushButtonXLStoXLSX.setEnabled(True)
            global_vars.ui.pushButtonProcessing.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False) 
            global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(False)             
            global_vars.ui.pushButtonConcat.setEnabled(False)  
            return                


        if not source_excels_list:
            global_vars.ui.project_folder_label.setStyleSheet('color: red')  
            global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 
       
            global_vars.ui.info_label.setStyleSheet('color: red')            
            global_vars.ui.info_label.setText('В папке проекта есть папка .Исходники/, но в ней нет файлов .xlsx')
                             
            self.warning_message = 'В папке проекта есть папка .Исходники/, но в ней нет файлов .xlsx'

            global_vars.ui.pushButtonXLStoXLSX.setEnabled(True)
            global_vars.ui.pushButtonProcessing.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False) 
            global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(False)             
            global_vars.ui.pushButtonConcat.setEnabled(False)  
            return    

        global_vars.ui.project_folder_label.setStyleSheet('color: green')  
        global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 

        global_vars.ui.info_label.setStyleSheet('color: blue')          
        global_vars.ui.info_label.setText('Запустите обработку')

        '''
        global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)
        global_vars.ui.pushButtonProcessing.setEnabled(True)
        global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(True)      
        global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(True)
        global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(True)          
        sleep(0.01)
        '''

   
                      
        print(f'run {self.message_title}')   

    def on_clicked(self):
        if os.path.exists('.session_folder'):
            with open('.session_folder', encoding='utf-8') as f:
                dir = f.readline()
        else:
            dir = ''
    
        global_vars.project_folder = QtWidgets.QFileDialog.getExistingDirectory(dir=dir) 


        if global_vars.project_folder:
            with open('.session_folder', 'w', encoding='utf-8') as f:
                    f.write(global_vars.project_folder)
        print('Запускаем поток  ')
        self.start() # Запускаем поток  
     
        
    def on_started(self): # Вызывается при запуске потока     
        print(f"on_started {self.message_title}")
        global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)
        global_vars.ui.pushButtonProcessing.setEnabled(False)
        global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)              
        global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False) 
        global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(False)              
        global_vars.ui.pushButtonConcat.setEnabled(False)    


    def on_finished(self): # Вызывается при завершении потока
        print(f"on_finished {self.error_message}")
        if self.error_message:
            QtWidgets.QMessageBox.critical(None,
                self.message_title,
                self.error_message,
                buttons=QtWidgets.QMessageBox.StandardButton.Ok)
            return 
        
        if self.warning_message:
            QtWidgets.QMessageBox.warning(None,
                self.message_title,
                self.warning_message,
                buttons=QtWidgets.QMessageBox.StandardButton.Ok)
            
            return    
        
        global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)
        global_vars.ui.pushButtonProcessing.setEnabled(True)
        global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(True)      
        global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(True)
        global_vars.ui.pushButtonDelChoosedMDFiles.setEnabled(True) 
 
