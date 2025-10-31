import pandas as pd
from pprint import pprint

df = pd.DataFrame(
    [
        {
            'a': 11,
            'b': 'bbb1'
        },
        {
            'a': 12,
            'b': 'bbb2'
        },
        {
            'a': 12,
            'b': 'bbb2'
        }
    ])

pprint(df)
df['a'] = df['a'].apply(str)
df['scep'] = df['a'] + '|' + df['b']
print(df)