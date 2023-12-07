import pandas as pd
import os

# create empty dataframe to store the combined data
combined_data = pd.DataFrame()
# speficy the directory the excel files are located
folder_path =  "Data"
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
all_numbers = pd.DataFrame({'編號': range(1, 53)})
result_wide = pd.merge(all_numbers, result_wide, on='編號', how='left')
result_wide.fillna(0, inplace=True)

# 兩個半套 ==============================================================
# Identify the index for '油條'
bread_index = result_wide[result_wide['編號'] == 34].index[0]

# Calculate the sum of corresponding values for '燒餅' and '兩個半套'
bread_sum = result_wide[result_wide['編號'].isin([34, 52])].sum()

# Update the values for '油條' row
result_wide.iloc[bread_index, 1:] = bread_sum.iloc[1:].values

# Identify the index for '燒餅'
roll_index = result_wide[result_wide['編號'] == 28].index[0]

# Calculate the sum of corresponding values for '燒餅' and '兩個半套'
roll_sum = result_wide[result_wide['編號'].isin([28, 52])].sum() + result_wide[result_wide['編號'].isin([52])].sum()

# Update the values for '燒餅' row
result_wide.iloc[roll_index, 1:] = roll_sum.iloc[1:].values

result_wide = result_wide.iloc[:-2]
# 兩個半套 ==============================================================

result_wide['名稱'] = ['冰炭香豆漿', '冰花生米漿', '冰炭香清漿', '冰豆米漿', '冰蜂蜜紅茶', '冰蜂蜜豆漿紅茶',
'熱炭香豆漿', '熱花生米漿', '熱炭香清漿', '熱豆米漿', '熱豆漿加蛋', '鹹豆漿', '鹹豆漿加蛋', '酸辣湯', '鹹飯糰', '甜飯糰',
'椒鹽飯糰', '素飯糰', '鹹飯糰夾蔥蛋', '素飯糰加散蛋', '椒鹽飯糰夾蔥蛋', '手工饅頭', '手工饅頭加蔥蛋', '手工鮮肉包',
'僧想吃菜包', '鹹酥餅', '麥芽甜餅', '燒餅', '燒餅夾蔥蛋', '燒餅油條一套', '燒餅油條夾蔥蛋', '燒餅油條夾酸菜', '燒餅夾牛肉',
'油條', '荷包蛋', '蔥花蛋', '馬來糕', '千層糕', '廣式蘿蔔糕', '廣式蘿蔔糕加蛋', '牛肉餡餅', '豬肉餡餅', '韭菜盒', 
'蘿蔔絲餅', '蘿蔔絲蛋餅', '小籠包', '小籠煎包', '家庭號豆漿', '家庭號清漿', '超辣辣椒醬2罐']

column_c = result_wide.pop('名稱')
result_wide.insert(1, '名稱', column_c)

result_daily = pd.DataFrame()

# 豆漿
soy_count = 1 * result_wide.loc[result_wide['編號'] == 1, '總計'].values[0] + \
               1 * result_wide.loc[result_wide['編號'] == 3, '總計'].values[0] + \
               0.5 * result_wide.loc[result_wide['編號'] == 4, '總計'].values[0] + \
               1 * result_wide.loc[result_wide['編號'] == 7, '總計'].values[0] + \
               1 * result_wide.loc[result_wide['編號'] == 9, '總計'].values[0] + \
               0.5 * result_wide.loc[result_wide['編號'] == 10, '總計'].values[0] + \
               0.5 * result_wide.loc[result_wide['編號'] == 11, '總計'].values[0] + \
               0.5 * result_wide.loc[result_wide['編號'] == 12, '總計'].values[0] + \
               0.5 * result_wide.loc[result_wide['編號'] == 13, '總計'].values[0] + \
               4 * result_wide.loc[result_wide['編號'] == 48, '總計'].values[0] + \
               4 * result_wide.loc[result_wide['編號'] == 49, '總計'].values[0]

soy = pd.DataFrame({'名稱': ['炭香豆漿'], 'Count': [soy_count]})

# 米漿
ricemilk_count = 1 * result_wide.loc[result_wide['編號'] == 2, '總計'].values[0] + \
                    0.5 * result_wide.loc[result_wide['編號'] == 4, '總計'].values[0] + \
                    1 * result_wide.loc[result_wide['編號'] == 8, '總計'].values[0] + \
                    0.5 * result_wide.loc[result_wide['編號'] == 10, '總計'].values[0]
                    
ricemilk = pd.DataFrame({'名稱': ['花生米漿'], 'Count': [ricemilk_count]})

# 紅茶
tea_count = 1 * result_wide.loc[result_wide['編號'] == 5, '總計'].values[0] + \
               1 * result_wide.loc[result_wide['編號'] == 6, '總計'].values[0]

tea = pd.DataFrame({'名稱': ['紅茶'], 'Count': [tea_count]})

# 酸辣濤
soup = pd.DataFrame({'名稱': ['酸辣湯'], 'Count': [result_wide.loc[result_wide['編號'] == 14, '總計'].values[0]]})

# 飯糰
riceball_count = result_wide.loc[result_wide['編號'].isin(range(15, 22)), '總計'].sum()
riceball = pd.DataFrame({'名稱': ['飯糰'], 'Count': [riceball_count]})

# 饅頭
bun_count = result_wide.loc[result_wide['編號'].isin(range(22, 24)), '總計'].sum()
bun = pd.DataFrame({'名稱': ['饅頭'], 'Count': [bun_count]})

# 鮮肉包
meatbun = pd.DataFrame({'名稱': ['鮮肉包'], 'Count': [result_wide.loc[result_wide['編號'] == 24, '總計'].values[0]]})

# 菜包
vegbun = pd.DataFrame({'名稱': ['菜包'], 'Count': [result_wide.loc[result_wide['編號'] == 25, '總計'].values[0]]})

# 鹹酥餅
meatcake = pd.DataFrame({'名稱': ['鹹酥餅'], 'Count': [result_wide.loc[result_wide['編號'] == 26, '總計'].values[0]]})

# 麥芽甜餅
maltcake = pd.DataFrame({'名稱': ['麥芽甜餅'], 'Count': [result_wide.loc[result_wide['編號'] == 27, '總計'].values[0]]})

# 燒餅
ovenroll_count = result_wide.loc[result_wide['編號'].isin(range(28, 34)), '總計'].sum()
ovenroll = pd.DataFrame({'名稱': ['燒餅'], 'Count': [ovenroll_count]})

# 油條
breadstick_count = result_wide.loc[result_wide['編號'].isin([30, 31, 32, 34]), '總計'].sum() + riceball_count
breadstick = pd.DataFrame({'名稱': ['油條'], 'Count': [breadstick_count]})

# 蛋
egg_count = result_wide.loc[result_wide['編號'].isin([11, 13, 19, 20, 21, 23, 29, 31, 35, 36, 40, 45]), '總計'].sum()
egg = pd.DataFrame({'名稱': ['蛋'], 'Count': [egg_count]})

# 牛肉片
beef = pd.DataFrame({'名稱': ['牛肉片'], 'Count': [result_wide.loc[result_wide['編號'] == 33, '總計'].values[0]]})

# 酸菜
pickle = pd.DataFrame({'名稱': ['酸菜'], 'Count': [result_wide.loc[result_wide['編號'] == 32, '總計'].values[0]]})

# 馬來糕
malay = pd.DataFrame({'名稱': ['馬來糕'], 'Count': [result_wide.loc[result_wide['編號'] == 37, '總計'].values[0]]})

# 千層糕
coldcake = pd.DataFrame({'名稱': ['千層糕'], 'Count': [result_wide.loc[result_wide['編號'] == 38, '總計'].values[0]]})

# 蘿蔔糕 count
turnipcake_count = result_wide.loc[result_wide['編號'].isin(range(39, 41)), '總計'].sum()
turnip = pd.DataFrame({'名稱': ['蘿蔔糕'], 'Count': [ovenroll_count]})

# 牛肉餡餅
beefcake = pd.DataFrame({'名稱': ['牛肉餡餅'], 'Count': [result_wide.loc[result_wide['編號'] == 41, '總計'].values[0]]})

# 豬肉餡餅
porkcake = pd.DataFrame({'名稱': ['豬肉餡餅'], 'Count': [result_wide.loc[result_wide['編號'] == 42, '總計'].values[0]]})

# 韭菜盒
chive = pd.DataFrame({'名稱': ['韭菜盒'], 'Count': [result_wide.loc[result_wide['編號'] == 43, '總計'].values[0]]})

# 蘿蔔絲餅
radish_count = result_wide.loc[result_wide['編號'].isin(range(44, 46)), '總計'].sum()
radish = pd.DataFrame({'名稱': ['蘿蔔絲餅'], 'Count': [radish_count]})

# 小籠包
xiaolongbao_count = result_wide.loc[result_wide['編號'].isin(range(46, 48)), '總計'].sum()
xiaolongbao = pd.DataFrame({'名稱': ['小籠包'], 'Count': [xiaolongbao_count]})

# 超辣辣椒醬
chili = pd.DataFrame({'名稱': ['超辣辣椒醬'], 'Count': [2 * (result_wide.loc[result_wide['編號'] == 50, '總計'].values[0])]})

# Concatenate the original DataFrame with the new row
result_daily = pd.concat([result_daily, soy, ricemilk, tea, soup, riceball, bun, meatbun, vegbun, meatcake, maltcake, 
                          ovenroll, breadstick, egg, beef, pickle, malay, coldcake, turnip, beefcake, porkcake, chive,
                          radish, xiaolongbao, chili], ignore_index=True)

# Export the hourly DataFrame to Excel
file_name = '銷售時段表.xlsx'
sheet_name = 'Sheet1'
result_wide.to_excel(file_name, sheet_name=sheet_name, index=False)

# Export the daily DataFrame to Excel
file_name2 = '每日明細.xlsx'
sheet_name2 = 'Sheet1'
result_daily.to_excel(file_name2, sheet_name=sheet_name2, index=False)