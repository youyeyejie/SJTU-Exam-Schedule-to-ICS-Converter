from datetime import datetime
import icalendar
import os
import sys
import pytz
import pysjtu
import json

# 获取当前执行脚本的终端路径
root_path = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(root_path, "config.json")

# 登录教务系统
if len(sys.argv) > 2:
    username = sys.argv[1]
    password = sys.argv[2]
elif os.path.exists(config_path):
    with open(config_path, "r") as f:
        config = json.load(f)
        username = config.get("jaccount")
        password = config.get("password")
    if not username or not password:
        username = input("请输入jaccount账号：")
        password = input("请输入jaccount密码：")
else:
    username = input("请输入jaccount账号：")
    password = input("请输入jaccount密码：")

print("正在登录教务系统...")
try:
    client = pysjtu.create_client(username, password)
except Exception as e:
    print("登录失败，请检查账号密码是否正确。")
    sys.exit(1)
print("登录成功。")

# 获取考试安排信息
if len(sys.argv) > 4:
    year = sys.argv[3]
    term = sys.argv[4]
    print("学年：", year)
    print("学期：", term)
else:
    year = input("请输入学年：")
    term = input("请输入学期[1，2，3]：")

year = int(year.split('-')[0])
term = int(term) - 1
exam_schedule = client.exam(year, term)
if not exam_schedule:
    print("未找到考试安排信息。")
    sys.exit(1)
else:
    print("成功获取考试安排信息。")
    for exam in exam_schedule:
        print(exam)
    print("正在生成ICS文件...")

# 生成ICS文件保存路径
ics_name = f"{year}-{year+1}-第{term+1}学期考试安排.ics"
directory = os.path.join(root_path, "data")
if not os.path.exists(directory):
    os.makedirs(directory)
ics_path = os.path.join(directory, ics_name)

# 创建日历对象
c = icalendar.Calendar()

# 设置日历的默认时区
c.timezone = pytz.timezone('Asia/Shanghai')

# 遍历考试安排信息来构建事件并添加到日历中
for exam in exam_schedule:
    course_name = exam.course_name
    course_id = exam.course_id
    location = exam.location
    date = exam.date
    time = exam.time
    # 设置时区为北京时间
    beijing_tz = pytz.timezone('Asia/Shanghai')

    # 组合日期和时间，转换为icalendar要求的格式
    start_datetime = datetime.combine(date, time[0], tzinfo=beijing_tz)
    end_datetime = datetime.combine(date, time[1], tzinfo=beijing_tz)

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