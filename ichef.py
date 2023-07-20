import pandas as pd
import os

# create empty dataframe to store the combined data
combined_data = pd.DataFrame()
# speficy the directory the excel files are located
folder_path = "data"
file_paths = [os.path.join(folder_path, file_name) for file_name in os.listdir(folder_path)]
excel_files = [file for file in file_paths if file.endswith('.xlsx')]
sorted_files = sorted(excel_files, key=lambda x: os.path.getmtime(x))

def remove_phrases_in_brackets(text):
    return pd.Series(text).str.replace(r'\s*\([^)]*\)', '', regex=True).values[0]

i = 0
for file_path in sorted_files:
    df = pd.read_excel(file_path)
    df['時段'] = str(i).zfill(2) + ":00-" + str(i+1).zfill(2) + ":00" 
    i = i + 1
    # Append the data to the all_data DataFrame
    combined_data =pd.concat([combined_data, df], ignore_index=True)

combined_data['號碼'] = combined_data['名稱'].apply(remove_phrases_in_brackets)
combined_data = combined_data[combined_data['號碼'].str.match(r'^\d')]
combined_data['號碼'] = combined_data['號碼'].str.extract(r'(\d+)')
result = combined_data.groupby(['號碼', '時段'])['銷售量'].sum().reset_index()
result_wide = result.pivot(index='號碼', columns='時段', values='銷售量').fillna(0)
result_wide['總計'] = result_wide.sum(axis=1)

result_wide['編號'] = result_wide.index.astype(int)
all_numbers = pd.DataFrame({'編號': range(1, 51)})
result_wide = pd.merge(all_numbers, result_wide, on='編號', how='left')
result_wide.fillna(0, inplace=True)

result_wide['名稱'] = ['冰炭香豆漿', '冰花生米漿', '冰炭香清漿', '冰豆米漿', '冰蜂蜜紅茶', '冰蜂蜜豆漿紅茶',
'熱炭香豆漿', '熱花生米漿', '熱炭香清漿', '熱豆米漿', '熱豆漿加蛋', '鹹豆漿', '鹹豆漿加蛋', '酸辣湯', '鹹飯糰', '甜飯糰',
'椒鹽飯糰', '素飯糰', '鹹飯糰夾蔥蛋', '素飯糰加散蛋', '椒鹽飯糰夾蔥蛋', '手工饅頭', '手工饅頭加蔥蛋', '手工鮮肉包',
'僧想吃菜包', '鹹酥餅', '麥芽甜餅', '燒餅', '燒餅夾蔥蛋', '燒餅油條一套', '燒餅油條夾蔥蛋', '燒餅油條夾酸菜', '燒餅夾牛肉',
'油條', '荷包蛋', '蔥花蛋', '馬來糕', '千層糕', '廣式蘿蔔糕', '廣式蘿蔔糕加蛋', '牛肉餡餅', '豬肉餡餅', '韭菜盒', 
'蘿蔔絲餅', '蘿蔔絲蛋餅', '小籠包', '小籠煎包', '家庭號豆漿', '家庭號清漿', '超辣辣椒醬2罐']

column_c = result_wide.pop('名稱')
result_wide.insert(1, '名稱', column_c)

# Export the DataFrame to Excel
file_name = '銷售時段表.xlsx'
sheet_name = 'Sheet1'

result_wide.to_excel(file_name, sheet_name=sheet_name, index=False)