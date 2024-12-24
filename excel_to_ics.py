import pandas as pd
from ics import Calendar, Event
from datetime import datetime
import os
import sys
import pytz

# 获取当前执行脚本的终端路径
running_path = os.getcwd()

# 检查是否提供了命令行参数
if len(sys.argv) > 1:
    file_path = sys.argv[1]
    ics_path = os.path.splitext(file_path)[0] + '.ics'
else:
    # 定义默认的文件路径和生成的ICS文件保存路径
    file_path = input("请输入文件路径（支持Excel或txt文件）：")
    file_path = os.path.join(running_path, file_path)
    ics_path = os.path.splitext(file_path)[0] + '.ics'

# 检查文件是否存在
if not os.path.exists(file_path):
    print("文件不存在，请检查文件路径是否正确。")
    sys.exit(1)

# 判断文件类型并读取文件
if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
    df = pd.read_excel(file_path)
elif file_path.endswith('.txt'):
    df = pd.read_csv(file_path, sep='\t', header=0)
else:
    print("不支持的文件格式。请提供Excel或txt文件。")
    sys.exit(1)

# 创建日历对象
c = Calendar()

# 设置日历的默认时区
c.timezone = pytz.timezone('Asia/Shanghai')

# 遍历DataFrame的每一行来构建事件并添加到日历中
for index, row in df.iterrows():
    course_name = row['课程名称']
    location = row['考试地点']
    # 获取考试时间字符串并处理
    time_str = row['考试时间']
    start_time, end_time = time_str[time_str.find("(")+1:time_str.find(")")].split("-")
    date_str = time_str[:time_str.find("(")].strip()
    # 将日期字符串转换为日期对象
    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
    # 设置时区为北京时间
    beijing_tz = pytz.timezone('Asia/Shanghai')
    # 组合日期和开始时间，转换为ics要求的格式
    start_datetime = datetime.strptime(date_str + " " + start_time.strip(), '%Y-%m-%d %H:%M')
    start_datetime = beijing_tz.localize(start_datetime)
    # 组合日期和结束时间，转换为ics要求的格式
    end_datetime = datetime.strptime(date_str + " " + end_time.strip(), '%Y-%m-%d %H:%M')
    end_datetime = beijing_tz.localize(end_datetime)

    event = Event()
    event.name = course_name + "考试"
    event.begin = start_datetime
    event.end = end_datetime
    event.location = location
    c.events.add(event)

# 将日历对象转换为.ics文件内容并写入文件
with open(ics_path, "w", encoding="utf-8") as f:
    f.writelines(c.serialize())

print(f"已生成ICS文件：{ics_path}")