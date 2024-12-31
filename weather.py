import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import re
import openpyxl

# 设置 matplotlib 后端
plt.switch_backend('TkAgg')

# 城市名称和对应的天气网页后缀
city_suffix = {
    '北京': '101010100',
    '上海': '101020100',
    '天津': '101030100',
    '重庆': '101040100',
    '哈尔滨': '101050101',
    '长春': '101060101',
    '沈阳': '101070101',
    '呼和浩特': '101080101',
    '石家庄': '101090101',
    '太原': '101100101',
    '济南': '101120101',
    '郑州': '101180101',
    '西安': '101110101',
    '兰州': '101160101',
    '银川': '101170101',
    '西宁': '101150101',
    '乌鲁木齐': '101130101',
    '合肥': '101220101',
    '南京': '101190101',
    '杭州': '101210101',
    '福州': '101230101',
    '南昌': '101240101',
    '长沙': '101250101',
    '武汉': '101200101',
    '成都': '101270101',
    '贵阳': '101260101',
    '昆明': '101290101',
    '南宁': '101300101',
    '拉萨': '101140101',
    '海口': '101310101',
    '广州': '101280101'
}

def parse_date_string(date_str):
    """
    解析非标准日期格式的日期字符串，例如 "23日（今天）"，"24日（明天）" 等,返回格式为 %Y-%m-%d 的日期字符串。
    """
    today = pd.Timestamp.today().date()

    if '今天' in date_str:
        return today.strftime('%Y-%m-%d')
    elif '明天' in date_str:
        return (today + pd.Timedelta(days=1)).strftime('%Y-%m-%d')
    elif '后天' in date_str:
        return (today + pd.Timedelta(days=2)).strftime('%Y-%m-%d')
    elif '周' in date_str:
        weekday = date_str[-2:-1]  # 提取周几的汉字
        days_ahead = {'一': 0, '二': 1, '三': 2, '四': 3, '五': 4, '六': 5, '日': 6}[weekday]
        target_date = today + pd.DateOffset(days=days_ahead)
        return target_date.strftime('%Y-%m-%d')
    else:
        # 默认处理，尝试提取日期部分
        match = re.search(r'\d{4}-\d{2}-\d{2}', date_str)
        if match:
            return match.group(0)
        else:
            raise ValueError(f"无法解析日期: {date_str}")

def get_weather_data(start_date, end_date, location):
    try:
        start_date = pd.to_datetime(start_date, format='%Y-%m-%d').strftime('%Y-%m-%d')
        end_date = pd.to_datetime(end_date, format='%Y-%m-%d').strftime('%Y-%m-%d')

        date_range = pd.date_range(start_date, end_date)
        weather_data = []

        if location in city_suffix:
            url_suffix = city_suffix[location]
            url = f"http://www.weather.com.cn/weather/{url_suffix}.shtml"
        else:
            print(f"未知地点: {location}")
            return pd.DataFrame(columns=['Date', 'Weather', 'Temperature', 'Wind'])

        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers)
        response.encoding = 'utf-8'

        soup = BeautifulSoup(response.text, 'html.parser')

        weather_container = soup.find('ul', class_='t clearfix')
        if not weather_container:
            raise ValueError('未能找到天气信息容器')

        weather_items = weather_container.find_all('li')

        for item in weather_items:
            try:
                date_str = item.find('h1').text.strip()
                date = parse_date_string(date_str)

                weather = item.find('p', class_='wea').text.strip()
                temperature_high = item.find('p', class_='tem').find('span')
                temperature_low = item.find('p', class_='tem').find('i').text.strip()
                wind = item.find('p', class_='win').find('span')['title']

                if temperature_high:
                    temperature_high = temperature_high.text.strip()
                else:
                    temperature_high = temperature_low

                temperature = f"{temperature_low}/{temperature_high}"

                # 将日期转换为统一格式进行比较
                current_date = pd.to_datetime(date, format='%Y-%m-%d').strftime('%Y-%m-%d')
                if current_date in date_range:
                    weather_data.append([date, weather, temperature, wind])

                # 如果当前日期超过了结束日期，就停止获取数据
                if current_date > end_date:
                    break
            except AttributeError as e:
                print(f"跳过无法解析的天气项：{e}")

        # 确保没有重复日期，并按日期排序
        weather_data = pd.DataFrame(weather_data, columns=['Date', 'Weather', 'Temperature', 'Wind'])
        weather_data = weather_data.drop_duplicates(subset=['Date']).sort_values(by='Date')

        # 补全缺失的日期并填充缺失数据
        weather_data = weather_data.set_index('Date').reindex(date_range.strftime('%Y-%m-%d')).ffill().reset_index()
        weather_data.columns = ['Date', 'Weather', 'Temperature', 'Wind']

        return weather_data

    except Exception as e:
        print(f"获取天气数据时出现错误：{e}")
        return pd.DataFrame(columns=['Date', 'Weather', 'Temperature', 'Wind'])

# 天气数据可视化应用程序
class WeatherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("天气数据统计可视化系统")
        self.root.geometry("800x800")  # 设置窗口大小

        # 设置默认字体
        self.zhfont1 = fm.FontProperties(fname='C:\\Windows\\Fonts\\simsun.ttc')

        # 时段选择组件
        self.start_date_label = tk.Label(root, text="开始日期：")
        self.start_date_label.place(x=10, y=10, width=300, height=30)
        self.start_date_entry = tk.Entry(root)
        self.start_date_entry.place(x=310, y=10, width=150, height=30)

        self.end_date_label = tk.Label(root, text="结束日期：")
        self.end_date_label.place(x=10, y=50, width=300, height=30)
        self.end_date_entry = tk.Entry(root)
        self.end_date_entry.place(x=310, y=50, width=150, height=30)

        self.location_label = tk.Label(root, text="地点：")
        self.location_label.place(x=110, y=90, width=100, height=30)
        self.location_entry = tk.Entry(root)
        self.location_entry.place(x=310, y=90, width=150, height=30)

        self.start_button = tk.Button(root, text="运行", command=self.start)
        self.start_button.place(x=200, y=130, width=200, height=30)

        # 表格和滚动条
        self.table_frame = tk.Frame(root)
        self.table_frame.place(x=10, y=180, width=780, height=200)

        self.table_scroll = ttk.Scrollbar(self.table_frame)
        self.table_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.table = ttk.Treeview(self.table_frame, yscrollcommand=self.table_scroll.set)
        self.table['columns'] = ('Date', 'Weather', 'Temperature', 'Wind')
        self.table.column('#0', width=0, stretch=tk.NO)  # 隐藏第一列
        for col in self.table['columns']:
            self.table.column(col, width=180)
            self.table.heading(col, text=col)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH)

        self.table_scroll.config(command=self.table.yview)

        # 图形绘制区域
        self.figure, (self.ax1, self.ax2) = plt.subplots(2, 1, figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().place(x=10, y=400, width=780, height=380)

        # 初始化时清空图形
        self.clear_plot()

    def clear_plot(self):
        self.ax1.clear()
        self.ax2.clear()
        self.ax1.set_xticks([])
        self.ax1.set_yticks([])
        self.ax2.set_xticks([])
        self.ax2.set_yticks([])
        self.canvas.draw()

    def start(self):
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()
        location = self.location_entry.get()

        try:
            start_date = pd.to_datetime(start_date_str, format='%Y-%m-%d').strftime('%Y-%m-%d')
            end_date = pd.to_datetime(end_date_str, format='%Y-%m-%d').strftime('%Y-%m-%d')

            df = get_weather_data(start_date, end_date, location)
            if not df.empty:
                self.show_data(df)
                self.visualize_data(df)
                self.save_to_excel(df, location)
            else:
                print("未能获取数据，请检查输入")
        except ValueError as e:
            print(f"日期格式错误：{e}")

    def show_data(self, df):
        self.table.delete(*self.table.get_children())
        for index, row in df.iterrows():
            self.table.insert("", "end", values=row.tolist())

    def visualize_data(self, df):
        self.ax1.clear()
        df['Temperature'] = df['Temperature'].apply(lambda x: int(x.split('/')[1].replace('℃', '')))
        df.plot(kind='line', x='Date', y='Temperature', ax=self.ax1, marker='o')

        # 设置字体属性
        self.ax1.set_title('Temperature Over Time', fontproperties=self.zhfont1, fontsize=12)
        self.ax1.set_xlabel('Date', fontproperties=self.zhfont1, fontsize=10)
        self.ax1.set_ylabel('Temperature (℃)', fontproperties=self.zhfont1, fontsize=10)

        # 设置刻度标签的字体属性
        for label in self.ax1.get_xticklabels() + self.ax1.get_yticklabels():
            label.set_fontproperties(self.zhfont1)

        # 绘制天气状况比例饼状图
        self.ax2.clear()
        weather_counts = df['Weather'].value_counts()
        weather_counts.plot(kind='pie', ax=self.ax2, autopct='%1.1f%%', startangle=90,
                            textprops={'fontproperties': self.zhfont1})

        self.ax2.set_ylabel('')  # 隐藏y轴标签
        #self.ax2.set_title('天气状况饼状图', fontproperties=self.zhfont1, fontsize=12)

        # 固定饼状图的位置在下方
        self.ax2.set_position([0.1, 0.1, 0.35, 0.35])  # 调整参数根据需要修改

        self.canvas.draw()

    def save_to_excel(self, df, location):
        # 数据预处理：使用前一天的数据填补缺失数据
        df = df.ffill()

        filename = f"{location}_weather_data.xlsx"
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Weather Data')
                workbook = writer.book
                worksheet = writer.sheets['Weather Data']

                for idx, col in enumerate(worksheet.columns, 1):
                    max_length = max(len(str(cell.value)) for cell in col)
                    worksheet.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = max_length

            print(f"天气数据已保存到文件: {filename}")
        except PermissionError:
            print(f"错误：没有写入 {filename} 的权限。请检查写入权限。")
        except Exception as e:
            print(f"保存到Excel时发生错误：{e}")

# 主程序入口
if __name__ == "__main__":
    root = tk.Tk()
    app = WeatherApp(root)
    root.mainloop()
