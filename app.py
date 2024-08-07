from flask import Flask, render_template, jsonify, request
from pymongo import MongoClient
import datetime
import matplotlib.pyplot as plt
import io
import base64
from matplotlib.font_manager import FontProperties
import plotly.graph_objs as go
import plotly.io as pio
from gtts import gTTS
# 使用 Agg 后端
import matplotlib
from pydub import AudioSegment
import os
import subprocess
matplotlib.use('Agg')
import numpy as np
app = Flask(__name__)
# 设置字体属性
# font_path = 'SimHei.ttf'  # 请确保这个路径正确，指向你的中文字体文件
# font_prop = FontProperties(fname=font_path)
plt.rcParams['font.family'] = "Microsoft JhengHei"
plt.rcParams['axes.unicode_minus']=False
# MongoDB configuration
client = MongoClient("mongodb://localhost:27017/")
db = client['PROD']
collection = db['INDEX_DATA']

def get_gauge(value, title):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={'text': title},
        gauge={
            'axis': {'range': [None, 4]},  # Adjust the range as needed
            'bar': {'color': "red", 'thickness': 0.95},
            'steps': [
                {'range': [0, 4], 'color': "lightblue"},],
            }))
    
    return pio.to_html(fig, full_html=False)


def get_image(fig):
    img = io.BytesIO()
    fig.savefig(img, format='png')
    img.seek(0)
    plot_url = base64.b64encode(img.getvalue()).decode('utf8')
    return plot_url
@app.route('/')
# def index():
#     return render_template('index.html')

# @app.route('/api/data', methods=['GET'])
def index():
    # Get today's date and the date 20 days ago
    today = datetime.datetime.today()
    twenty_days_ago = today - datetime.timedelta(days=20)
    
    # Convert dates to strings if necessary for comparison with MongoDB data
    today_str = today.strftime('%Y-%m-%d')
    twenty_days_ago_str = twenty_days_ago.strftime('%Y-%m-%d')
    
    # Fetch data from MongoDB collection and filter for the last 20 days
    data = list(collection.find())[-30:]
    data_all = list(collection.find())
    # Ensure data is not None or empty
    if not data:
        data = []

    # 提取 '選擇權未平倉總和(億)' 列
    total_oi_option_values = [item['選擇權未平倉總和(億)'] for item in data_all]
    total_oi_percentiles = np.percentile(total_oi_option_values, [10, 90])
    bottom10Percent = total_oi_percentiles[0]
    top10Percent = total_oi_percentiles[1]

    total_oi_option_valuessum = [item['選擇權未平倉總和增減(億)'] for item in data_all]
    total_oi_percentilessum = np.percentile(total_oi_option_valuessum, [10, 90])
    bottom10Percent_opsum = total_oi_percentilessum[0]
    top10Percent_opsum = total_oi_percentilessum[1]

    dates = [item['日期'] for item in data]
    closing_prices = [item['收盤指數'] for item in data]
    foreign_investment = [item['外資(億元)'] for item in data]
    margin_changes = [item['融資增減(億元)'] for item in data]
    short_changes = [item['融券增減(張)'] / 1000 for item in data]
    futures_long = [item['外資期貨未平倉多方(口)'] for item in data]
    futures_short = [item['外資期貨未平倉空方(口)'] for item in data]
    options_buy = [item['外資選擇權買權未平倉(億)'] for item in data]
    options_sell = [-item['外資選擇權賣權未平倉(億)'] for item in data]  # 添加负号


    latest_record = collection.find().sort("日期", -1).limit(1)[0]
    latest_long = latest_record['多']
    latest_short = latest_record['空']
    print(latest_long, latest_short)
    gauge_long = get_gauge(latest_long, "多指標")
    gauge_short = get_gauge(latest_short, "空指標")
    
    #語音合成
    latest_close = round(latest_record['收盤指數'],2)
    latest_foreign = round(latest_record['外資(億元)'],2)
    latest_margin = round(latest_record['融資增減(億元)'],3)
    latest_margin_short = int(latest_record['融券增減(張)'])
    latest_futures_long = int(latest_record['外資期貨未平倉多方(口)'])
    latest_futures_short = int(latest_record['外資期貨未平倉空方(口)'])
    latest_options_buy = round(latest_record['外資選擇權買權未平倉(億)'],3)
    latest_options_sell = round(latest_record['外資選擇權賣權未平倉(億)'],3)
    


    # Plot 1: Foreign Investment
    fig1, ax1 = plt.subplots()
    ax1.plot(dates, closing_prices, color='blue', marker='o', label='加權指數')
    ax2 = ax1.twinx()
    bars = ax2.bar(dates, foreign_investment, alpha=0.6, label='外資買賣超', 
                   color=['red' if val > 0 else 'green' for val in foreign_investment])
    fig1.autofmt_xdate(rotation=45)
    plt.setp(ax1.get_xticklabels(), fontsize=8)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('加權指數')
    ax2.set_ylabel('外資買賣超')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plot1 = get_image(fig1)
    plt.close(fig1)  # Close the figure after saving to avoid display issues

    # Plot 2: Margin Changes and Short Changes
    fig2, ax1 = plt.subplots()
    ax1.plot(dates, closing_prices, color='blue', marker='o', label='加權指數')
    ax2 = ax1.twinx()
    width = 0.4  # Width of the bars
    bars_margin = ax2.bar([d - width/2 for d in range(len(dates))], margin_changes, width=width, alpha=0.6, label='融資變化(億)', color='red', align='center')
    ax3 = ax1.twinx()
    ax3.spines['right'].set_position(('outward', 60))
    bars_short = ax3.bar([d + width/2 for d in range(len(dates))], short_changes, width=width, alpha=0.6, label='融券變化(千張)', color='green', align='center')

    # 确保零轴对齐
    max_y2 = max(max(margin_changes), abs(min(margin_changes)))
    max_y3 = max(max(short_changes), abs(min(short_changes)))
    max_y = max(max_y2, max_y3)

    ax2.set_ylim(-max_y, max_y)
    ax3.set_ylim(-max_y, max_y)

    fig2.autofmt_xdate(rotation=45)
    plt.setp(ax1.get_xticklabels(), fontsize=8)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('加權指數')
    ax2.set_ylabel('融資變化')
    ax3.set_ylabel('融券變化(千口)')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper center')
    ax3.legend(loc='upper right')
    plot2 = get_image(fig2)
    plt.close(fig2)  # Close the figure after saving to avoid display issues
    # Plot 3: Futures Positions
    fig3, ax1 = plt.subplots()
    ax1.plot(dates, closing_prices, color='blue', marker='o', label='加權指數')
    ax2 = ax1.twinx()
    width = 0.4  # Width of the bars
    ax2.bar([d - width/2 for d in range(len(dates))], futures_long, width=width, alpha=0.6, label='外資期貨未平倉多方(口)', color='red', align='center')
    ax2.bar([d + width/2 for d in range(len(dates))], futures_short, width=width, alpha=0.6, label='外資期貨未平倉空方(口)', color='green', align='center')
    fig3.autofmt_xdate(rotation=45)
    plt.setp(ax1.get_xticklabels(), fontsize=8)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('加權指數')
    ax2.set_ylabel('外資期貨部位')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plot3 = get_image(fig3)
    plt.close(fig3)  # Close the figure after saving to avoid display issues

    # Plot 4: Options Positions
    fig4, ax1 = plt.subplots()
    ax1.plot(dates, closing_prices, color='blue', marker='o', label='加權指數')
    ax2 = ax1.twinx()
    ax2.bar(dates, options_buy, alpha=0.6, label='外資選擇權買權未平倉(億)', color='red')
    ax2.bar(dates, options_sell, alpha=0.6, label='外資選擇權賣權未平倉(億)', color='green')
    fig4.autofmt_xdate(rotation=45)
    plt.setp(ax1.get_xticklabels(), fontsize=8)
    ax1.set_xlabel('日期')
    ax1.set_ylabel('加權指數')
    ax2.set_ylabel('外資選擇權')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plot4 = get_image(fig4)
    plt.close(fig4)  # Close the figure after saving to avoid display issues

    # 生成语音文件
    speech_text = f"這是{today_str}的台股資訊，昨日收盤價為{latest_close}點,外資買賣超{latest_foreign}億,外資期貨多方未平倉{latest_futures_long}口,外資期貨多方未平倉{latest_futures_short}口,外資選擇權多方未平倉{latest_options_buy}億,外資選擇權空方未平倉{latest_options_sell}億,融資增減{latest_margin}億,融券增減{latest_margin_short}張,多指標為{latest_long},空指標為{latest_short}。"
    tts = gTTS(speech_text, lang='zh')
    audio_file = 'C:/Users/FinLab615_82/Future_Analysis/static/speech.mp3'
    tts.save(audio_file)
    if not os.path.exists(audio_file):
        raise FileNotFoundError(f"Audio file not found: {audio_file}")

    return render_template('index.html', today_date=today_str, data=data, plot1=plot1, plot2=plot2, plot3=plot3, plot4=plot4, gauge_long=gauge_long, gauge_short=gauge_short, bottom10Percent=bottom10Percent, top10Percent=top10Percent, bottom10Percent_opsum=bottom10Percent_opsum, top10Percent_opsum=top10Percent_opsum)

if __name__ == '__main__':
    app.run(debug=True,host='0.0.0.0', port=5002)