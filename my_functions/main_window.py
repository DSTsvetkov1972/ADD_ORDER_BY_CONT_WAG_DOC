import pandas as pd

import datetime
#import pyperclip
from datetime import datetime
from colorama import init, Fore, Back, Style
from PySide6 import QtWidgets, QtCore
import global_vars
from my_functions.checks import header_checker

def checkHeaders(headers):
    i = 0
    sep = "~"
    while i< len(headers):
        header = headers[i]
        if sep in header:
            sep +="~"
            i = 0
        else:
            i += 1
    headers = ['NoName' if header == '' else header for header in headers]
    i = len(headers)-1
    while i != 0:
        header = headers[i]
        if header in headers[:i]:
            headers[i] = headers[i] + sep + str(headers.count(header)-1)
            continue
        else:
            i -= 1      
    headers = ['Строка в исходнике~0' if header == 'Строка в исходнике' else header for header in headers]
    #headers = ['~№' if header == '№' else header for header in headers]    
    return headers

def translit(kir):
    translation = kir.maketrans({'а':'a',
                           'б':'b',
                           'в':'v',
                           'г':'g',
                           'д':'d',
                           'е':'e',
                           'ё':'yo',
                           'ж':'zh',
                           'з':'z',
                           'и':'i',
                           'й':'j',
                           'к':'k',
                           'л':'l',
                           'м':'m',
                           'н':'n',
                           'о':'o',
                           'п':'p',
                           'р':'r',
                           'с':'s',
                           'т':'t',
                           'у':'u',
                           'ф':'f',
                           'х':'h',
                           'ц':'ts',
                           'ч':'ch',
                           'ш':'sh',
                           'щ':'shh',
                           'ъ':'',
                           'ы':'y',
                           'ь':'',
                           'э':'e',
                           'ю':'yu',
                           'я':'ya',
                           ' ':'_',
                           '.':'_'})
    return kir.lower().translate(translation)

"""
def load_file_sheet_name():
        MyComboBoxFormats.set_eanbled_all(False)
        if ('loaded_file' in global_vars.__dict__.keys() and 
        'loaded_sheet_name'  in global_vars.__dict__.keys()):
            if (global_vars.file != global_vars.loaded_file or
                global_vars.sheet_name != global_vars.loaded_sheet_name):
                
                print(Fore.CYAN,'Загружаем лист целиком (A)',Fore.WHITE)

                #global_vars.header_row = 0
                global_vars.ui.footer_label.setStyleSheet('color: blue')

                if global_vars.file[-4:] in ['.xls', 'xlsx', 'xlsm', 'xlsb', '.ods']:
                    global_vars.ui.footer_label.setText(f'Загружаем весь лист "{global_vars.ui.comboSheets.currentText()}"...')
                    global_vars.df = pd.read_excel(global_vars.file, header = None, dtype= 'string', engine = 'calamine', sheet_name = global_vars.sheet_name)
                else:
                    global_vars.ui.footer_label.setText(f'Загружаем весь файл...')
                    global_vars.df = from_file_to_csv(global_vars.file)
                global_vars.df.fillna('', inplace = True)
                global_vars.df.index += 1
                global_vars.df.can_load_file = True
                global_vars.loaded_file = global_vars.file
                global_vars.loaded_sheet_name = global_vars.sheet_name  
                global_vars.ui.footer_label.setStyleSheet('color: green')
                global_vars.ui.footer_label.setText(f'Весь лист "{global_vars.ui.comboSheets.currentText()}" загружен!')
                 
        else:

            print(Fore.CYAN,'Загружаем лист целиком (B)',Fore.WHITE)

            #global_vars.header_row = 0 
            if global_vars.file[-4:] in ['xlsx','xlsm','.xls','.ods']:
                global_vars.df = pd.read_excel(global_vars.file, header = None, dtype= 'string', engine = 'calamine', sheet_name = global_vars.sheet_name)
            else:
                global_vars.df = from_file_to_csv(global_vars.file)
            global_vars.df.fillna('', inplace = True)
            global_vars.df.index += 1
            global_vars.df.can_load_file = True
            global_vars.loaded_file = global_vars.file
            global_vars.loaded_sheet_name = global_vars.sheet_name 
            global_vars.ui.footer_label.setStyleSheet('color: green')
            global_vars.ui.footer_label.setText(f'Весь лист "{global_vars.ui.comboSheets.currentText()}" загружен!')
             
        MyComboBoxFormats.set_eanbled_all(True)
        print(Fore.CYAN,'Загружаем лист целиком (C)',Fore.WHITE)
"""

def from_file_to_csv(f):
    with open(f, encoding='utf-8') as fn:
        all_lines = fn.readlines()
        max = 0
        min = 0
        for l in all_lines:
            l_list = l.split('\t')
            if len(l_list) > max:
                max = len(l_list)
            if len(l_list) < min:
                min = len(l_list)
        all_lines_added = [] 
        for l in all_lines:
            l_list = l.split('\t')
            all_lines_added.append(l + '\t'*(max - len(l_list)))

    list_to_df = [row.split('\t') for row in all_lines_added]     
     
    df =  pd.DataFrame(list_to_df) 
    df.fillna('', inplace= True) 
    df.map(str)        

    return df

def fill_in_view_table (df_view):

    row_number = 0
    for row in df_view.itertuples():

        column_number = 0
        for item in row[1:]:
            cellinfo = QtWidgets.QTableWidgetItem(str(item))
            cellinfo.setFlags(
                QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled
            )
            global_vars.ui.tableWidget.setItem(row_number, column_number, cellinfo)
            column_number += 1
        row_number +=1    

def fill_in_err_table(err_columns_list:list, header_errors:list):
    global_vars.ui.err_tableWidget.setRowCount(len(header_errors))
    global_vars.ui.err_tableWidget.setColumnCount(len(err_columns_list)) 
    global_vars.ui.err_tableWidget.setHorizontalHeaderLabels(err_columns_list)
    row_number = 0
    for header_err_tuple in header_errors:

        column_number = 0
        for item in header_err_tuple:
            cellinfo = QtWidgets.QTableWidgetItem(str(item))
            cellinfo.setFlags(
                QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled
            )
            global_vars.ui.err_tableWidget.setItem(row_number, column_number, cellinfo)
            column_number += 1
        row_number +=1    
    global_vars.ui.err_tableWidget.resizeColumnsToContents()

def fill_in_table():
    global_vars.ui.footer_text.setVisible(False)
    df_view = global_vars.df.iloc[global_vars.header_row : global_vars.header_row + 16]

    global_vars.ui.tableWidget.setRowCount(len(df_view))
    global_vars.ui.tableWidget.setColumnCount(len(df_view.columns))    

    #if global_vars.header_row == 0:
    #    global_vars.horizontal_headers = checkHeaders(list(map(str,df_view.columns)))
    #else:
    #    global_vars.horizontal_headers = checkHeaders(global_vars.df.iloc[global_vars.header_row-1])

    if global_vars.header_row == 0:
        global_vars.horizontal_headers = (list(map(str,df_view.columns)))
    else:
        global_vars.horizontal_headers = (global_vars.df.iloc[global_vars.header_row-1])        

 

    global_vars.ui.tableWidget.setHorizontalHeaderLabels(global_vars.horizontal_headers)
    global_vars.ui.tableWidget.setVerticalHeaderLabels(map(str,list(range(global_vars.header_row+1, global_vars.header_row + 17))))
    # verticalHeaders = ['Формат >>'] + list(map(str,list(df_view.index)))
    #global_vars.ui.tableWidget.setVerticalHeaderLabels(verticalHeaders)

    column_number = 0

    global_vars.column_formats = []
    MyComboBoxFormats.instances = []
    MyComboBoxFormats.all_err_df = pd.DataFrame(columns = ['column_number','Сообщение','Ячейка LN','Ячейка RNCN','Значение'])

    fill_in_view_table(df_view)
    header_errors_list = header_checker(global_vars.sample_columns, list(global_vars.horizontal_headers))

    if header_errors_list:
        fill_in_err_table(['колонка', 'ошибка'], header_errors_list)
        global_vars.ui.err_tableWidget.setVisible(True)
        global_vars.ui.pushButtonPreprocessing.setEnabled(False)
    else:
        global_vars.ui.pushButtonPreprocessing.setEnabled(True)

def header_down(self):
    global_vars.ui.footer_text.setVisible(False)
    global_vars.ui.pushButtonLoader.setEnabled(False)
    global_vars.ui.err_tableWidget.setVisible(False)
    if global_vars.header_row > 1:
        global_vars.header_row -= 1

        if global_vars.header_row == 0: 
            global_vars.ui.footer_label.setStyleSheet('color: red')
            global_vars.ui.footer_label.setText(f"Достигнуто начало файла!")
        else:   
            global_vars.ui.footer_label.setStyleSheet('color: green')
            global_vars.ui.footer_label.setText(f"Заголовки опущены! Заголовки в строке {global_vars.header_row}")
    else:
        global_vars.ui.footer_label.setStyleSheet('color: red')
        global_vars.ui.footer_label.setText("Ниже опустить не возможно, заголовки!")
    fill_in_table()
    '''
    header_errors_list = header_checker(global_vars.sample_columns, global_vars.horizontal_headers)
    if header_errors_list:
        fill_in_err_table(['колонка', 'ошибка'], header_errors_list)
        global_vars.ui.err_tableWidget.setVisible(True)
    '''

def convert_str_to_date(s):
    if s == '1970-01-01 03:00:00':
        return ""
    else: 
        return datetime.strptime(s, "%Y-%m-%d").date()


if __name__ == '__main__':
    print(convert_str_to_date('03.05.2024  0:00:00'))
    print(convert_str_to_date('1970-01-01 03:00:00'))