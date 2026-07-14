"""
Скрипт для обработки файла csv с заявками абитуриентов через госуслуги от Кристины Си
"""
import pandas as pd



df = pd.read_csv('data/report_spo_2026_2.csv',delimiter=';',encoding='cp1251')

print(df.shape)

df = df[df['region_name'] == 'Республика Бурятия']
print(df.shape)


svod_df = pd.pivot_table(df,['order_id'],
                         index=['short_title'],
                         aggfunc='count',margins=True,
                         margins_name='Итого')

svod_df = svod_df.reset_index()
svod_df.columns = ['ПОО','Количество заявлений']

svod_df.to_excel('data/Свод по большому файлу Госуслуг.xlsx',index=False)


print('Lindy Booth')



