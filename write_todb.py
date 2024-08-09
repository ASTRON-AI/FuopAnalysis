import pandas as pd
from pymongo import MongoClient

import pandas as pd
file_path = './大盤檢查表20240617.xlsm' #EXCEL名稱跟路徑
df = pd.read_excel(file_path, sheet_name='歷史資料區')  # 替换为你的文件路径和工作表名称
columns = [
    "日期",
    "收盤指數", "漲跌", "成交量(億元)","比較","五日成交量平均(億元)", "外資(億元)", "投信(億元)", "自營(億元)", "融資(億元)", "融資增減(億元)", "融券(張張)", "融券增減(張)", 
    "外資期貨未平倉多方(口)","外資期貨未平倉空方(口)","外資期貨未平倉淨單(口)", "外資期貨未平倉增減(口)", "外資選擇權買權未平倉(億)", "外資選擇權賣權未平倉(億)", 
    "自營商選擇權買權未平倉(億)", "自營商選擇權賣權未平倉(億)", "選擇權未平倉總和(億)", "選擇權未平倉總和增減(億)", 
    "選擇權未平倉總和增減累計(億)", "多", "空"
    ]
print(len(columns))
new_df = df.iloc[1:,:]
new_df.columns = columns
new_df['日期'] = new_df['日期'].dt.strftime('%Y-%m-%d')
# 连接到MongoDB
client = MongoClient('mongodb://localhost:27017/')  # 修改为你的MongoDB连接字符串
db = client['PROD']  # 替换为你的数据库名称
collection = db['INDEX_DATA']  # 替换为你的集合名称

# 将DataFrame转换为字典并插入到MongoDB
data_dict = new_df.to_dict("records")
collection.insert_many(data_dict)

print("Excel data has been successfully stored to MongoDB.")