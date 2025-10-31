import  pandas as pd
import os
from datetime import datetime 
from colorama import Fore
from collections import Counter
import sqlite3
import global_vars

def init_project():
    if not os.path.exists(os.path.join(global_vars.project_folder, '.Обработка')):
        os.mkdir(os.path.join(global_vars.project_folder, '.Обработка'))
    
    if not os.path.exists(os.path.join(global_vars.project_folder, '.Размеченные')):
        os.mkdir(os.path.join(global_vars.project_folder, '.Размеченные'))

    if not os.path.exists(os.path.join(global_vars.project_folder, '.Файлы для 1-С')):
        os.mkdir(os.path.join(global_vars.project_folder, '.Файлы для 1-С'))    

    conn = sqlite3.connect(os.path.join(global_vars.project_folder, "files_info.db"))
    with conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS src_files_info (file TEXT, modifyed_time TEXT)")  
        # cur.execute("CREATE TABLE IF NOT EXISTS actual_src_files_info (file TEXT, modifyed_time TEXT)")  
        cur.execute("CREATE TABLE IF NOT EXISTS md_files_info (file TEXT, modifyed_time TEXT)")  
        # cur.execute("CREATE TABLE IF NOT EXISTS actual_md_files_info (file TEXT, modifyed_time TEXT)")                          

def value_searcher(col, value):
    #found = col[col==value]
    found = col.apply(lambda x: value in str(x))
    found = found[found]
    if found.size == 0:
        return "-"
    elif (len(found)) > 1:
        return "несколько"
    else:
        return str(found.index[0]+1)
    
def headers_checker(header_rows_df):
    i = 0
    errors_list = []
    for column in header_rows_df.columns:
        i += 1
        if i >= 3:
            if not pd.isna(header_rows_df[column].iloc[0]) and not pd.isna(header_rows_df[column].iloc[1]):
                errors_list.append(str(i))
    if errors_list != []:
        return 'Колонки: ' + ', '.join(errors_list) + ' промаркированы и "с заполнением" и "без"' 

    for column in header_rows_df.columns:
        i += 1
        if i >= 3:
            if not pd.isna(header_rows_df[column].iloc[0]) and not pd.isna(header_rows_df[column].iloc[1]):
                errors_list.append(str(i))
    if errors_list != []:
        return 'Колонки: ' + ', '.join(errors_list) + ' промаркированы и "с заполнением" и "без"'   
    
def repeating_headers_checker(header_rows_df):

    headers_0 = [i for i in header_rows_df.loc[0][2:] if pd.notna(i)]
    headers_1 = [i for i in header_rows_df.loc[1][2:] if pd.notna(i)]
    if headers_0 == [] and headers_1 == []:
        return ("Не выбрано ни одного заголовка")    

    errors_list = [f"{k}-{v}" for k,v in Counter(headers_0 + headers_1).items() if v>1]
    if errors_list != []:
        return ("Повторяющиеся заголовки: " + ", ".join(errors_list))

def marking_checker(sheet_rem, s, f, header_rows):
    errors_list = []

    if str(sheet_rem) != 'nan':
        return sheet_rem

    if s == "-" and f == "-":
        return "-"
    elif s == "-" and f != "-":
        errors_list.append('Маркер s не задан')
    elif s != "-" and f == "-":
        errors_list.append('Маркер f не задан')        

    if s == "несколько":
        errors_list.append("Маркер s проставлен в нескольких строках")
    elif s != "-":
        if int(s) < 3:
            errors_list.append('Маркер s расположен выше области таблицы')        
            
    if f == "несколько":
        errors_list.append("Маркер f проставлен в нескольких строках")
    elif f != "-":
        if int(f) < 2:
            errors_list.append('Маркер f расположен выше области таблицы') 

    if (s != "несколько" and f != "несколько" and
        s !="-" and f != "-" and
        int(s) > int(f)):
        errors_list.append('Маркер f расположен выше маркера s')  

    headers_errors =  headers_checker(header_rows)
    if headers_errors:
        errors_list.append(headers_errors) 

    repeating_headers_errors = repeating_headers_checker(header_rows)
    if repeating_headers_errors:
        errors_list.append(repeating_headers_errors)       

    if errors_list:        
        return ("; " + "\n").join(errors_list)
    else:
        return "ok"

    

def refresh_files_info(folder):
    if folder=='.Размеченные':
        table='md_files_info'
    elif folder=='.Исходники': 
        table='src_files_info'

    folder_path = os.path.join(global_vars.project_folder, folder)

    files = list(os.walk(folder_path))[0][2]    
    files = [(file, f"{os.path.getmtime(os.path.join(folder_path , file))}") for file in files if file[0] != "~"] 

    conn = sqlite3.connect(os.path.join(global_vars.project_folder, "files_info.db"))
    with conn:
        actual_files_info_df = pd.DataFrame(files, columns=['file','modifyed_time'])
        actual_files_info_df.to_sql(table, conn, index=False, if_exists='replace')
        

def check_files_modified(folder):
    print(f'check_files_modified {folder}')
    if folder=='.Размеченные':
        table='md_files_info'
    elif folder=='.Исходники': 
        table='src_files_info'

    folder_path = os.path.join(global_vars.project_folder, folder)  

    files = list(os.walk(folder_path))[0][2]    
    files = [(file, f"{os.path.getmtime(os.path.join(folder_path, file))}") for file in files if file[0] != "~"]

    conn = sqlite3.connect(os.path.join(global_vars.project_folder, "files_info.db"))
    with conn:
        cur = conn.cursor()


        actual_md_files_info_df = pd.DataFrame(files, columns=['file','modifyed_time'])
        actual_md_files_info_df.to_sql(f'actual_{table}', conn, index=False, if_exists='replace')

        files_in_table = list(cur.execute(f"SELECT * FROM {table}"))
        if  not files_in_table:
            print(Fore.RED, "Таблица БД не содержит записей", files_in_table, Fore.WHITE)
            return True
        
        cur.execute(f"SELECT t.file FROM {table} AS t INNER JOIN actual_{table} AS at ON t.file = at.file WHERE t.modifyed_time <> at.modifyed_time")
        
        files_modified = [file[0] for file in cur.fetchall()]
        #input('Ждем ввод')
        if  files_modified:
            print(Fore.RED, f"Файлы были пересохранены {files_modified}", Fore.WHITE)            
            return files_modified
        
        new_files = list(cur.execute(f"SELECT * FROM actual_{table} AS at LEFT JOIN {table} AS t ON t.file = at.file WHERE t.file IS NULL"))
        if  new_files:
            print(Fore.RED, f"Появились новые файлы {new_files}", Fore.RESET)           
            return True 

        deleted_files = list(cur.execute(f"SELECT * FROM {table} AS t LEFT JOIN actual_{table} AS at ON t.file = at.file WHERE at.file IS NULL"))
        if  deleted_files:
            print(Fore.RED, f"Некоторые файлы были удалены {deleted_files}", Fore.RESET)           
            return True 
    
        return False
        

def max_column(file, sheet_name):
    df = pd.read_excel(file, sheet_name = sheet_name)
    return (len(df.columns))

def clean_process_folder(project_folder):
    """
    Удаляет из папкпи .Обработка файлы, которые можно удалить
    """

    if not project_folder:
        return 
    
    try:
        processing_files = list(os.walk(os.path.join(project_folder,'.Обработка')))[0][2]
    except:
        print(Fore.RED, "Папка .Обработка не была очищенна. Или её нет или папка проекта не выбрана", Fore.RESET)
        return
    
    for pr_file_number, pr_file in enumerate(processing_files, 1):
        try:
            os.remove(os.path.join(project_folder, '.Обработка', pr_file))
        except PermissionError:
            print(Fore.RED, f"{pr_file_number} из {len(processing_files)} Файл {pr_file} не может быть удален из .Обработка", Fore.RESET)
            pass
