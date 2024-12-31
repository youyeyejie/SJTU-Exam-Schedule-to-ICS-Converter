import pandas as pd
from datetime import datetime
import icalendar
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
c = icalendar.Calendar()

# 设置日历的默认时区
c.timezone = pytz.timezone('Asia/Shanghai')

# 遍历DataFrame的每一行来构建事件并添加到日历中
for index, row in df.iterrows():
    course_id = row['课程代码']
    course_name = row['课程名称']
    location = row['考试地点']
    time_str = row['考试时间']

    # 解析考试时间字符串
    start_time, end_time = time_str[time_str.find("(")+1:time_str.find(")")].split("-")
    date_str = time_str[:time_str.find("(")].strip()

    # 将日期字符串转换为日期对象
    date = datetime.strptime(date_str, '%Y-%m-%d')
    start_time = datetime.strptime(start_time, '%H:%M').time()
    end_time = datetime.strptime(end_time, '%H:%M').time()

    # 设置时区为北京时间
    beijing_tz = pytz.timezone('Asia/Shanghai')
    # 组合日期和时间，转换为icalendar要求的格式
    start_datetime = datetime.combine(date, start_time, tzinfo=beijing_tz)
    end_datetime = datetime.combine(date, end_time, tzinfo=beijing_tz)

    # 创建事件对象
    event = icalendar.Event()
    event.add("summary", f"{course_id} {course_name} 考试")
    event.add("location", location)
    event.add("dtstart", start_datetime)
    event.add("dtend", end_datetime)
    c.add_component(event)

# 将日历对象转换为.ics文件内容并写入文件
with open(ics_path, "wb") as f:
    icalendar_with_seperate_lines = c.to_ical().replace(b"END:VEVENT", b"END:VEVENT\n")
    f.write(icalendar_with_seperate_lines)

print(f"已生成ICS文件：{ics_path}")