import pandas as pd
import os
import win32api, win32con
import global_vars

import datetime
import pyperclip
from clickhouse_driver import Client
from datetime import datetime
# from params import *
from cryptography.fernet import Fernet
from colorama import init, Fore, Back, Style
import re 
from PySide6 import QtWidgets, QtCore
import global_vars
# from my_widgets.my_combo_box_formats import MyComboBoxFormats


def get_df_of_click(query: str):
    try:
        params = get_params()
        connection=Client(
            host=params[0],
            port=params[1],
            database=params[2],
            user=params[3],
            password=params[4],
            secure=True,verify=False
            )
        
        with connection:
            return connection.query_dataframe(query)
        
    except Exception as e:
        global_vars.ui.login_label.setStyleSheet("color: red")
        global_vars.ui.login_label.setText(f"Ошибка DWH. Описание ошибки скопировано в буфер обмена {str(e)}")
        pyperclip.copy(e)
        
def execute_sql_click(query, operation_name = ''):
        print(Fore.CYAN, query, Fore.WHITE)
        try: 
            thread_start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(Fore.YELLOW + f'Запущен процесс: {operation_name} - {thread_start_time}' + Fore.WHITE)  
            params = get_params()
            connection=Client(host   = params[0],
                            port     = params[1],
                            database = params[2],
                            user     = params[3],
                            password = params[4],
                            secure=True,verify=False)
            connection.execute(query)
            thread_finish_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(Fore.GREEN + f'Закончен процесс: {operation_name} (начали {thread_start_time}, закончили {thread_finish_time})' + Fore.WHITE)#, выполнялся {(thread_finish_time-thread_start_time).total_seconds()} секунд' + Fore.WHITE)
            return True
        except Exception as e:
            print(Fore.RED + f'Авария при выполнении: {operation_name}.\nОшибка:\n{e}' + Fore.WHITE)
            return False

def insert_from_df(dwh_table_name, df, operation_name):
        #try: 
            thread_start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(Fore.YELLOW + f'Запущен процесс: {operation_name} - {thread_start_time}' + Fore.WHITE)  
            params = get_params()
            connection=Client(host   = params[0],
                            port     = params[1],
                            database = params[2],
                            user     = params[3],
                            password = params[4],
                            secure=True,
                            verify=False,
                            settings={'use_numpy': True})
            
            print(Fore.GREEN, df, Fore.WHITE)

            connection.insert_dataframe(f'INSERT INTO {dwh_table_name} VALUES', df)

            thread_finish_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            print(Fore.GREEN + f'Закончен процесс для: {operation_name} (начали {thread_start_time}, закончили {thread_finish_time})' + Fore.WHITE)#, выполнялся {(thread_finish_time-thread_start_time).total_seconds()} секунд' + Fore.WHITE)
        #except Exception as e:
        #    print(Fore.RED + f'Авария при выполнении для: {operation_name}.\nОшибка:\n{e}' + Fore.WHITE)
        #    pass        
        
def get_params():   
        if not os.path.exists(os.path.join('.config')): 
            global_vars.ui.login_label.setStyleSheet("color: red")
            global_vars.ui.login_label.setText("Подключитесь к DWH!")
            return ['','','','','']
        with open(os.path.join('.config')) as config_file:
            params = config_file.read()
        decoded_text = Fernet(b'lXgjsyWLG2R-nAWC1vBkz-FWFzeWFi-71rNMiO2ON40=').decrypt(params).decode('utf-8')
        return(decoded_text.split('\n'))
 
def get_last_version(label_get_new_version):
    sql = """
        SELECT 
            max(version) new_version,
            argMax(message,version) new_version_message
        FROM
            (SELECT 
                toInt64OrNull(
                    replace(
                        splitByChar('|',log_info )[1],
                        'version',
                        ''
                    )
                ) as version,
                splitByChar('|',log_info )[2] AS message
            FROM 
                audit._check_your_file 
            WHERE 
                log_info LIKE '%version%')
        """

    last_version_info = get_df_of_click(sql)
    #return(last_version_info)
    last_version_number  = last_version_info['new_version'][0]
    last_version_message = last_version_info['new_version_message'][0]

    if last_version_number > version:
        label_get_new_version.config(text = 'Версия %s доступна для скачивания'%last_version_number)
        label_get_new_version.bind('<Button-1>', lambda x:show_message(last_version_message))
    else:
        label_get_new_version.config(text = '')
        label_get_new_version.unbind('<Button-1>')


def connection_settings_file_creator(CLICK_HOST, CLICK_PORT, CLICK_DBNAME, CLICK_USER, CLICK_PWD):
    try: 
        params = ('%s\n%s\n%s\n%s\n%s')%(CLICK_HOST,CLICK_PORT,CLICK_DBNAME,CLICK_USER,CLICK_PWD)
        config_file = os.path.join('.config') 
        encoded_text = Fernet(b'lXgjsyWLG2R-nAWC1vBkz-FWFzeWFi-71rNMiO2ON40=').encrypt(params.encode('utf-8'))           
        if os.path.isfile(config_file):
            win32api.SetFileAttributes(config_file, win32con.FILE_ATTRIBUTE_NORMAL)


        with open (config_file,'wb') as file:         
            file.write(encoded_text)

        win32api.SetFileAttributes(config_file, win32con.FILE_ATTRIBUTE_HIDDEN)
              
        """  
        messagebox.showinfo(MSG_BOX_TITLE, 'Удалось подключиться к DWH!')
        filemenu = tkinter.Menu(mainmenu, tearoff=0)
        filemenu.add_command(label="Проверить соединение", command = lambda: Log_in_check(root,mainmenu))
        filemenu.add_command(label="Сменить пользователя", command = lambda: Log_in(root,mainmenu))
        filemenu.add_command(label="Выйти", command = lambda: Log_out(root,mainmenu))
        mainmenu.delete(3)
        mainmenu.add_cascade(label=get_params()[3],menu = filemenu, foreground = 'green') 
        """
    except Exception as e:
        print(f'Не удалось подключиться к DWH!\n{e}\nПроверьте параметры и повторите попрытку!')


def log_out():
    if os.path.exists(os.path.join('.config')):
        os.remove(os.path.join('.config')) 
        global_vars.ui.action_log_in_check.setEnabled(False)         
        global_vars.ui.action_log_out.setEnabled(False) 
        global_vars.ui.login_label.setStyleSheet("color: red")
        global_vars.ui.login_label.setText("Вы отключились от DWH!")
        global_vars.log_in_status = False       



    
