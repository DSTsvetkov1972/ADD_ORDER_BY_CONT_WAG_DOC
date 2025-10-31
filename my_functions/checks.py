import pandas as pd
import re
import time 

no_errs_df = pd.DataFrame(columns = ['column_number','Сообщение','Ячейка LN','Ячейка RNCN','Значение'])

def date_parser(date_str):
        patterns = ['%d-%m-%Y %H:%M:%S',
                    '%d-%m-%Y %H:%M',
                    '%d-%m-%Y',
                    '%d.%m.%Y %H:%M:%S',
                    '%d.%m.%Y %H:%M',
                    '%d.%m.%Y'
                    '%d.%m.%Y %H:%M:%S',
                    '%d.%m.%Y %H:%M',
                    '%d.%m.%Y',
                    '%Y-%m-%d %H:%M:%S.%f',                    
                    '%Y-%m-%d %H:%M:%S',
                    '%Y-%m-%d %H:%M',
                    '%Y-%m-%d'                 
                    ]
                    
        for pattern in patterns:
            try:
                dt = time.strptime(date_str, pattern)
                result = f"{dt.tm_year}-{str(dt.tm_mon).zfill(2)}-{str(dt.tm_mday).zfill(2)} {str(dt.tm_hour).zfill(2)}:{str(dt.tm_min).zfill(2)}:{str(dt.tm_sec).zfill(2)}"
                return result
            except:
                pass
        return ''

def preprocess_datetime(df):
    df['Дата отправки (проверенная)'] = df['Дата отправки'].apply(date_parser)
    err_df = df[['Дата отправки','Дата отправки (проверенная)']][df['Дата отправки (проверенная)']=='']
    err_list = [('Дата отправки', err[0], err[1]) for err in err_df.itertuples()]
    return err_list

def container_parser(container_number):
        container_number = container_number.upper().replace('К', 'K').replace('Т','T').replace(' ','')
        pattern = re.compile(r'[A-Z]{4}\d{7}')

        if len(container_number) != 11 or re.match(pattern, container_number) is None:
            return ''
        else:
            return container_number

def preprocess_container(df):
    df['№ контейнера (проверенный)'] = df['№ контейнера'].apply(container_parser)
    err_df = df[['№ контейнера','№ контейнера (проверенный)']][df['№ контейнера (проверенный)']=='']
    err_list = [('№ контейнера', err[0], err[1]) for err in err_df.itertuples()]
    return err_list

def header_checker(sample_columns:list, header_row:list )->list:
    """функция проверяет, чтобы таблица содержала нужные заголовки и
    чтобы заголовки не повторялись.
    Возвращает список недостающих и повторяющихся заголовков
    """

    err_list = []

    for column in sample_columns:
        if column not in header_row:
            err_list.append((column,'колонка отсутствует'))
        elif header_row.count(column) > 1:
            err_list.append((column,f'колонок с таким названием: {header_row.count(column)}'))


    return err_list    

if __name__ == '__main__':
    df = pd.DataFrame(['a','b','2022-12-03','c', '01-01-2022'], columns=['Дата отправки'])
    print(preprocess_datetime(df))
