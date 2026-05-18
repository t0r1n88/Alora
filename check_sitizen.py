"""
Скрипт для проверки выгрузки гражданства
"""

import pandas as pd


df = pd.read_excel('data/Билет в будушее сводка.xlsx',dtype=str)

print(df.columns)
print(df.shape)
df = df[df['Архивация'] == 'Нет']
print(df.shape)




