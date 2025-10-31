from PySide6 import QtWidgets, QtCore
from colorama import Fore
import global_vars 
import os, shutil
import win32com.client as win32
from openpyxl import Workbook,load_workbook,styles
from openpyxl.utils.cell import get_column_letter
from my_threads.functions import max_column
from copy import copy
from time import sleep

class XLS_TO_xlsxThread(QtCore.QThread):
 
    mysignal = QtCore.Signal(str)


    def __init__ (self, parent=None):
        QtCore.QThread.__init__(self, parent) 
        self.message_title = "Обработка"
     
        
    def convert_xls_to_xlsx(self, file_to_convert):
        file_converted = '.'.join(file_to_convert.split('.')[:-1])+'.xlsx'
        
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            wb = excel.Workbooks.Open(file_to_convert)
            wb.SaveAs(file_converted, FileFormat = 51)
            excel.Quit()
            os.remove(file_to_convert)
            return file_converted

        except Exception as e:
            print(e)
            return False
        finally:
            # RELEASES RESOURCES
            wb = None
            excel = None


    def on_start_convert_xls_to_xlsx(self, file_to_convert):

        print(Fore.MAGENTA, 'on_start_convert_xls_to_xlsx', file_to_convert, Fore.RESET)
        if file_to_convert[-4:] != '.xls' and file_to_convert[-4:] != 'xlsm':
            return False
        else:
            home_folder = os.path.expanduser('~')
            base_name = os.path.basename(file_to_convert)
            
            file_to_convert_in_home = os.path.join(home_folder, base_name)

            print(file_to_convert_in_home)
            if os.path.isfile(file_to_convert_in_home):
                os.remove(file_to_convert_in_home)    
                
            shutil.move(file_to_convert, home_folder) # os.path.join(home_folder, os.path.basename(file_to_convert)))     
  
            file_converted = os.path.join(home_folder, base_name[:-4] + '.xlsx')   
            
            print(file_converted)
            if os.path.isfile(file_converted):
                os.remove(file_converted)         

        print(Fore.MAGENTA, 'on_start_convert_xls_to_xlsx', file_to_convert_in_home, Fore.RESET)
        return file_to_convert_in_home
          
    def run_convert_xls_to_xlsx(self, file_to_convert):
        print(Fore.YELLOW, file_to_convert, Fore.RESET )
        file_to_convert_in_home = self.on_start_convert_xls_to_xlsx(file_to_convert)
        if file_to_convert_in_home:
            file_converted = self.convert_xls_to_xlsx(file_to_convert_in_home)
            if file_converted:
                shutil.move(file_converted, os.path.join(global_vars.project_folder, '.Исходники'))
            

    def run(self):
        self.error_message = ""
        self.warning_message = ""
        
        print(list(os.walk(global_vars.project_folder)))
        xls_files = list(os.walk(os.path.join(global_vars.project_folder,'.Исходники')))[0][2]
#        xls_files = [file for file in xls_files if file[-4:]=='.xls']
        xls_files = [file for file in xls_files if file[-4:]=='.xls' or file[-4:]=='xlsm']
        print(xls_files)
        print(Fore.YELLOW, f'run {global_vars.project_folder}',Fore.RESET)      
        
        for i, xls_file in enumerate(xls_files, 1):
            global_vars.ui.info_label.setStyleSheet('color: blue')             
            global_vars.ui.info_label.setText(f"Конвертируем .xls=>.xlsx {i} из {len(xls_files)}. {xls_file}.")
            print(f"Конвертируем {i} из {len(xls_files)}. {xls_file}.")
               
            if os.path.exists(os.path.join(global_vars.project_folder, '.Исходники', xls_file[:-4]+'.xlsx')):
                self.error_message = f"В папке .Исходники есть файлы {xls_file[:-4]}.xls и {xls_file[:-4]}.xlsx. Один из них надо удалить или переимновать!"
                break
            try:
                self.run_convert_xls_to_xlsx(os.path.join(global_vars.project_folder, '.Исходники', xls_file)) 
            except Exception as e:
                self.error_message = f"{xls_file}.xls\n{e}"
                break
       




    def on_clicked(self):
        #global_vars.project_folder = QtWidgets.QFileDialog.getExistingDirectory()        
        self.start() # Запускаем поток  
        print(Fore.BLUE, f'on_clicked {global_vars.project_folder}',Fore.RESET)        


    def on_started(self): # Вызывается при запуске потока
        global_vars.ui.pushButtonChooseProjectFolder.setEnabled(False) 
        global_vars.ui.pushButtonXLStoXLSX.setEnabled(False)               
        global_vars.ui.pushButtonProcessing.setEnabled(False)
        global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)        
        global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False)
        global_vars.ui.pushButtonConcat.setEnabled(False)        


    def on_finished(self): # Вызывается при завершении потока
        sleep(0.1)

        if self.error_message:
            print(self.error_message) 
            global_vars.ui.info_label.setStyleSheet('color: red')             
            global_vars.ui.info_label.setText(self.error_message.replace('\n',' '))
            QtWidgets.QMessageBox.critical(None,
                                           self.message_title,
                                           self.error_message,
                                           buttons=QtWidgets.QMessageBox.StandardButton.Ok)
            global_vars.ui.pushButtonChooseProjectFolder.setEnabled(True)
            global_vars.ui.pushButtonXLStoXLSX.setEnabled(True)   
            global_vars.ui.pushButtonProcessing.setEnabled(False)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(False)        
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(False)
        elif self.warning_message:
            global_vars.ui.info_label.setStyleSheet('color: red')             
            global_vars.ui.info_label.setText(self.warning_message.replace('\n',' '))
            QtWidgets.QMessageBox.warning(None,
                                           self.message_title,
                                           self.warning_message,
                                           buttons=QtWidgets.QMessageBox.StandardButton.Ok)             
        else:
            global_vars.ui.project_folder_label.setStyleSheet('color: green')  
            global_vars.ui.project_folder_label.setText(f'Папка проекта: {global_vars.project_folder}') 
            global_vars.ui.info_label.setStyleSheet('color: green')             
            global_vars.ui.info_label.setText('Файлы успешно конвертированы из .xls в .xlsx')
            
            global_vars.ui.pushButtonChooseProjectFolder.setEnabled(True)
            global_vars.ui.pushButtonXLStoXLSX.setEnabled(True)   
            global_vars.ui.pushButtonProcessing.setEnabled(True)
            global_vars.ui.pushButtonOpenChoosedFiles.setEnabled(True)        
            global_vars.ui.pushButtonOpenChoosedMDFiles.setEnabled(True)
