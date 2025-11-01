import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.DataFrame(
    [
        {
            'a': 11,
            'aa': 11,
            'b': 'bbb1'
        },
        {
            'a': 12,
            'aa': 12,            
            'b': 'bbb2'
        },
        {
            'a': 12,
            'aa': 12,            
            'b': 'bbb3'
        }
    ])

rows = dataframe_to_rows(df)

print(list(rows)[2:])