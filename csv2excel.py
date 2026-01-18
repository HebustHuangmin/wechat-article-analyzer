import pandas as pd
df = pd.read_csv('结果.csv', encoding='utf-8')
df.to_excel('结果.xlsx', index=False, engine='openpyxl')
print('OK')
