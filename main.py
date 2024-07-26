import pandas as pd
import requests
import pymongo
import numpy as np
from datetime import datetime, timedelta
def get_futDailyMarketExcel():
    # 目标URL
    url = "https://www.taifex.com.tw/cht/3/futDailyMarketExcel"
    # 发送请求获取网页内容
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            print(df.head())
        else:
            print("未能找到任何表格。")
    else:
        print("请求失败，状态码：", response.status_code)
    
def get_futContractsDateExcel():
    url = "https://www.taifex.com.tw/cht/3/futContractsDateExcel"
    # 发送请求获取网页内容
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib') 
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[1]  # 假设我们需要第一个表格
            # styled_df = df.style
            # display(styled_df)
            外資未平倉多方 = float(df.iloc[2,9]) + float(df.iloc[11,9]) /4
            外資未平倉空方 = float(df.iloc[2,11]) + float(df.iloc[11,11]) /4
            print("外資未平倉多方：", 外資未平倉多方)
            print("外資未平倉空方：", 外資未平倉空方)
            外資期貨未平倉淨單 = 外資未平倉多方 - 外資未平倉空方
            return 外資未平倉多方, 外資未平倉空方, 外資期貨未平倉淨單
        else:
            print("未能找到任何表格。")
            return None, None
    else:
        print("请求失败，状态码：", response.status_code)
        return None, None

def get_callsAndPutsDateExcel():
    url = "https://www.taifex.com.tw/cht/3/callsAndPutsDateExcel"
    response = requests.get(url)
    if response.status_code == 200:
        tables = pd.read_html(response.text, flavor='html5lib')
        if len(tables) > 0:
            df = tables[2]  # 假设我们需要第一个表格
            外資選擇權買權未平倉, 外資選擇權賣權未平倉 = df.iloc[2, 15] * 1000 / 10 ** 8, df.iloc[5, 15] * 1000 / 10 ** 8
            自營商選擇權買權未平倉, 自營商選擇權賣權未平倉 = df.iloc[0, 15] * 1000 / 10 ** 8 , df.iloc[3, 15] * 1000 / 10 ** 8
            print("外資選擇權買權未平倉：", 外資選擇權買權未平倉)
            print("外資選擇權賣權未平倉：", 外資選擇權賣權未平倉)
            print("自營商選擇權買權未平倉：", 自營商選擇權買權未平倉)
            print("自營商選擇權賣權未平倉：", 自營商選擇權賣權未平倉)
            選擇權未平倉總和 = 外資選擇權買權未平倉 - 外資選擇權賣權未平倉 + 自營商選擇權買權未平倉 - 自營商選擇權賣權未平倉
            print("選擇權未平倉總和：", 選擇權未平倉總和)
            return 外資選擇權買權未平倉, 外資選擇權賣權未平倉, 自營商選擇權買權未平倉, 自營商選擇權賣權未平倉, 選擇權未平倉總和
        else:
            print("未能找到任何表格。")
            return None, None, None, None, None
    else:
        print("请求失败，状态码：", response.status_code)
        return None, None, None, None, None
def get_time():
    url = "https://www.taifex.com.tw/cht/3/largeTraderFutQryTbl"
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            print(df.iloc[0,0])
            print(type(df.iloc[0,0]))
            return df.iloc[0,0]
        else:
            print("未能找到任何表格。")
    else:
        print("请求失败，状态码：", response.status_code)
def get_largeTraderFutQryTbl():
    url = "https://www.taifex.com.tw/cht/3/largeTraderFutQryTbl"
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[1]  # 假设我们需要第一个表格
            print(df.columns)
        else:
            print("未能找到任何表格。")
    else:
        print("请求失败，状态码：", response.status_code)

def get_threepeoplebuysell():
    url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=html"
    # 发送请求获取网页内容
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            自營商 , 投信 , 外資 = df.iloc[0,3] / 10 ** 8, df.iloc[2,3] / 10 ** 8, df.iloc[3,3] / 10 ** 8
            print("自營商：", 自營商)
            print("投信：", 投信)
            print("外資：", 外資)
            return 外資, 投信, 自營商
        else:
            print("未能找到任何表格。")
            return None, None, None
    else:
        print("请求失败，状态码：", response.status_code)
        return None, None, None

def get_indexcloseprice():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&type=IND"
    response = requests.get(url)
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            df.set_index(df.columns[0], inplace=True)
            收盤指數 = df.loc['發行量加權股價指數',df.columns[0]] # 取得發行量加權股價指數的收盤價
            return 收盤指數
        else:
            print("未能找到任何表格。")
            return None
    else:
        print("请求失败，状态码：", response.status_code)
        return None
def get_MRMARGIN():
    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN?response=html&selectType=MS&date="
    response = requests.get(url)

    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            融資餘額 , 融券餘額 = df.iloc[2,5] * 1000 / 10 ** 8, df.iloc[1,5]
            return 融資餘額, 融券餘額
        else:
            print("未能找到任何表格。")
            return None, None
    else:
        print("请求失败，状态码：", response.status_code)
        return None, None

def get_indexvolume():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=html&date=&type=MS"
    response = requests.get(url)
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            df.set_index(df.columns[0], inplace=True)
            index_volume = df.loc['1.一般股票',df.columns[0]] / 10 ** 8
            return index_volume
        else:
            print("未能找到任何表格。")
            return None
    else:
        print("请求失败，状态码：", response.status_code)
        return None
def get_NIGHTFUTURE():
    # 目标URL
    url = "https://www.taifex.com.tw/cht/3/futContractsDateAh"
    # 发送请求获取网页内容
    response = requests.get(url)
    # 检查请求是否成功
    if response.status_code == 200:
        # 使用pandas读取HTML表格
        tables = pd.read_html(response.text, flavor='html5lib')
        
        # 检查是否成功读取到表格
        if len(tables) > 0:
            df = tables[0]  # 假设我们需要第一个表格
            外資淨額 = df.iloc[2,7] + df.iloc[8,7] /4 
            print("外資淨額：", 外資淨額)
            return 外資淨額
        else:
            print("未能找到任何表格。")
            return None
    else:
        print("请求失败，状态码：", response.status_code)
        return None
def get_history_data():
    # 连接到MongoDB
    client = pymongo.MongoClient('mongodb://localhost:27017/')  # 修改为你的MongoDB连接字符串
    db = client['PROD']  # 替换为你的数据库名称
    collection = db['INDEX_DATA']  # 替换为你的集合名称

    # 从MongoDB集合中读取数据
    data = list(collection.find())

    # 将数据转换为DataFrame
    df = pd.DataFrame(data)

    # 移除默认的MongoDB ObjectId字段（如果存在）
    if '_id' in df.columns:
        df.drop(columns=['_id'], inplace=True)

    return df
def simulate():
    data = {
    "收盤指數": np.random.randint(15000, 16000, size=5),
    "漲跌": np.random.randint(-100, 100, size=5),
    "成交量": np.random.randint(100000, 200000, size=5),
    "五日成交量平均": [0]*5,  # To be calculated
    "外資": np.random.randint(-500, 500, size=5),
    "投信": np.random.randint(-50, 50, size=5),
    "自營": np.random.randint(-100, 100, size=5),
    "融資": np.random.randint(1000, 2000, size=5),
    "融資增減": np.random.randint(-200, 200, size=5),
    "融券": np.random.randint(500, 1000, size=5),
    "融券增減": np.random.randint(-100, 100, size=5),
    "外資期貨未平倉多方": np.random.randint(1000, 2000, size=5),
    "外資期貨未平倉空方": np.random.randint(500, 1500, size=5),
    "外資期貨未平倉增減": np.random.randint(-100, 100, size=5),
    "外資選擇權買權未平倉": np.random.rand(5),
    "外資選擇權賣權未平倉": np.random.rand(5),
    "自營商選擇權買權未平倉": np.random.rand(5),
    "自營商選擇權賣權未平倉": np.random.rand(5),
    "選擇權未平倉總和": np.random.randint(100, 500, size=5),
    "選擇權未平倉總和增減": np.random.randint(-50, 50, size=5),
    "選擇權未平倉總和增減累計": [0]*5,  # To be calculated
    "多": np.random.randint(50, 150, size=5),
    "空": np.random.randint(50, 150, size=5)
    }

    # Create the DataFrame
    df = pd.DataFrame(data)
    return df
def main():
    columns = [
    "日期",
    "收盤指數", "漲跌", "成交量(億元)","比較","五日成交量平均(億元)", "外資(億元)", "投信(億元)", "自營(億元)", "融資(億元)", "融資增減(億元)", "融券(張張)", "融券增減(張)", 
    "外資期貨未平倉多方(口)","外資期貨未平倉空方(口)","外資期貨未平倉增減(口)", "外資期貨未平倉增減(口)", "外資選擇權買權未平倉(億)", "外資選擇權賣權未平倉(億)", 
    "自營商選擇權買權未平倉(億)", "自營商選擇權賣權未平倉(億)", "選擇權未平倉總和(億)", "選擇權未平倉總和增減(億)", 
    "選擇權未平倉總和增減累計(億)", "多", "空"
    ]
    df = pd.DataFrame(columns=columns)
    # 判断当前时间
    current_hour = datetime.now().hour
    
    if current_hour == 7:  # 如果是早上7点，则日期设为昨天
        日期 = (datetime.now() - timedelta(1)).strftime("%Y-%m-%d")
    else:  # 否则日期设为今天
        日期 = datetime.now().strftime("%Y-%m-%d")
    #日期 = get_time()
    # 删除同日期的记录
    # 连接到MongoDB
    client = pymongo.MongoClient('mongodb://localhost:27017/')  # 修改为你的MongoDB连接字符串
    db = client['PROD']  # 替换为你的数据库名称
    collection = db['INDEX_DATA']  # 替换为你的集合名
    collection.delete_many({"日期": 日期})
    收盤指數 = get_indexcloseprice()
    大盤成交量 = get_indexvolume()
    外資, 投信, 自營商 = get_threepeoplebuysell()
    融資, 融券 = get_MRMARGIN()
    外資期貨未平倉多方, 外資期貨未平倉空方 ,外資期貨未平倉淨單= get_futContractsDateExcel()
    外資選擇權買權未平倉, 外資選擇權賣權未平倉, 自營商選擇權買權未平倉, 自營商選擇權賣權未平倉, 選擇權未平倉總和 = get_callsAndPutsDateExcel()
    歷史資料 = get_history_data()
    前日收盤 = 歷史資料.iloc[-1, 1]
    漲跌 = 收盤指數 - 前日收盤
    五日成交量平均 =(sum(歷史資料.iloc[-4:, 3])+大盤成交量) / 5
    比較 = ">" if 大盤成交量 > 五日成交量平均 else "<"
    前日融資, 前日融券 = 歷史資料.iloc[-1, 9], 歷史資料.iloc[-1, 11]
    前日外資期貨未平倉淨單 = 歷史資料.iloc[-1, 15]
    外資期貨未平倉增減 = 外資期貨未平倉淨單 - 前日外資期貨未平倉淨單
    前日選擇權未平倉總和 = 歷史資料.iloc[-1, 21]
    今日選擇權未平倉總和增減 = 選擇權未平倉總和 - 前日選擇權未平倉總和
    前日選擇權未平倉總和累積 = 歷史資料.iloc[-1, 23]
    今日選擇權未平倉總和增減累積 = 前日選擇權未平倉總和累積 + 今日選擇權未平倉總和增減
    percentile_v = np.percentile(歷史資料['選擇權未平倉總和(億)'], 90)
    percentile_w = np.percentile(歷史資料['選擇權未平倉總和增減(億)'], 90)
    多 = (
        (外資選擇權買權未平倉 >= 1).astype(int) +
        (選擇權未平倉總和 >= percentile_v).astype(int) +
        (今日選擇權未平倉總和增減 >= percentile_w).astype(int)
    )
    percentile_v = np.percentile(歷史資料['選擇權未平倉總和(億)'], 10)
    percentile_w = np.percentile(歷史資料['選擇權未平倉總和增減(億)'], 10)
    空 = (
    (外資選擇權賣權未平倉 >= 1).astype(int) +
    (選擇權未平倉總和 <= percentile_v).astype(int) +
    (今日選擇權未平倉總和增減 <= percentile_w).astype(int)
    )
    夜盤期貨變動 = get_NIGHTFUTURE()

    data = {
        "日期": 日期, "收盤指數": 收盤指數, "漲跌": 漲跌, "成交量(億元)": 大盤成交量, "比較": 比較, "五日成交量平均(億元)": 五日成交量平均, 
        "外資(億元)": 外資, "投信(億元)": 投信, "自營(億元)": 自營商, "融資(億元)": 融資, "融資增減(億元)": 融資 - 前日融資, 
        "融券(張張)": 融券, "融券增減(張)": 融券 - 前日融券, "外資期貨未平倉多方(口)": 外資期貨未平倉多方, "外資期貨未平倉空方(口)": 外資期貨未平倉空方, 
        "外資期貨未平倉淨單(口)": 外資期貨未平倉淨單, "外資期貨未平倉增減(口)": 外資期貨未平倉增減, "外資選擇權買權未平倉(億)": 外資選擇權買權未平倉, 
        "外資選擇權賣權未平倉(億)": 外資選擇權賣權未平倉, "自營商選擇權買權未平倉(億)": 自營商選擇權買權未平倉, "自營商選擇權賣權未平倉(億)": 自營商選擇權賣權未平倉, 
        "選擇權未平倉總和(億)": 選擇權未平倉總和, "選擇權未平倉總和增減(億)": 今日選擇權未平倉總和增減, "選擇權未平倉總和增減累計(億)": 今日選擇權未平倉總和增減累積, 
        "多": 多, "空": 空, "夜盤期貨變動":夜盤期貨變動
    }
    print(len(data))
    # 将所有numpy数据类型转换为原生Python数据类型
    data = {k: (int(v) if isinstance(v, (np.integer, np.int64)) else float(v) if isinstance(v, (np.float32, np.float64)) else v) for k, v in data.items()}
    
    # 插入数据
    collection.insert_one(data)

    print("Data inserted into MongoDB:", data)
    
if __name__ == "__main__":
    main()