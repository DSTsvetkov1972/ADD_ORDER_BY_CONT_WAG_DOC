from PySide6 import QtCore, QtWidgets
from colorama import Fore
from my_functions.dwh import get_params, Client
from my_functions.main_window import translit, load_file_sheet_name
import global_vars
import pandas as pd
import re, os

class LogInCheck(QtCore.QThread):
    def __init__ (self, parent=None):
        QtCore.QThread.__init__(self, parent)


    def run(self):
        if not os.path.exists(os.path.join('.config')): 
            global_vars.ui.action_log_in_check.setEnabled(False)
            global_vars.ui.action_log_out.setEnabled(False)            
            global_vars.ui.login_label.setStyleSheet("color: red")
            global_vars.ui.login_label.setText("Подключитесь к DWH!")
            global_vars.log_in_status = False
            return 
        else:
            params = get_params()

            global_vars.ui.login_label.setStyleSheet("color: blue")
            global_vars.ui.login_label.setText(f"Проверяем подключение пользователя {params[3]} к базе {params[2]}")
            try:
                connection=Client(host     = params[0],
                                port     = int(params[1]),
                                database = params[2],
                                user     = params[3],
                                password = params[4],
                                secure   = True,
                                verify   = False) 
                try:
                    connection.execute(f"""CREATE OR REPLACE TABLE {params[2]}.load_any_table_2_connection_test
                            ENGINE = MergeTree()
                            ORDER BY test AS SELECT 1 AS test
                            """)
                    connection.execute(f"""DROP TABLE IF EXISTS {params[2]}.load_any_table_2_connection_test""")

                    global_vars.ui.login_label.setStyleSheet("color: green")
                    global_vars.ui.login_label.setText(f"Пользователь {params[3]} подключен к базе {params[2]}") 
                    global_vars.ui.action_log_in_check.setEnabled(True)
                    global_vars.ui.action_log_out.setEnabled(True) 
                    global_vars.log_in_status = True                
                    return  
                except:
                    connection.execute("""SELECT 1""")   
                    global_vars.ui.login_label.setStyleSheet("color: red")
                    global_vars.ui.login_label.setText(f"Пользователь {params[3]} подключен к базе {params[2]}, НО БЕЗ ПРАВА СОЗДАВАТЬ ТАБЛИЦЫ!!!") 
                    global_vars.ui.action_log_in_check.setEnabled(True)
                    global_vars.ui.action_log_out.setEnabled(True)  
                    global_vars.log_in_status = False 
            except:
                global_vars.ui.login_label.setStyleSheet("color: red")
                global_vars.ui.login_label.setText(f"Не можем подключить {params[3]} к базе {params[2]}, проверьте параметры подключения. Возможно не работает DWH, нет интернета или...")                              
                global_vars.log_in_status = False

         



    def starter(self):
        pass

    def on_started(self): # Вызывается при запуске потока
        pass

    def on_finished(self): # Вызывается при завершении потока
        pass
