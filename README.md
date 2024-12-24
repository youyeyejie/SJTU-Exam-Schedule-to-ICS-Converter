# Exam Schedule to ICS Converter

该脚本用于将 SJTU 教学信息服务网导出的考试安排从 Excel 文件转换为 ICS 文件，以便导入到日历应用中。

## 使用方法

### 依赖安装

在运行脚本之前，请确保已安装以下依赖项：

```bash
pip install pandas ics pytz
```

### 运行脚本

你可以通过以下两种方式运行脚本：

1. 使用命令行参数指定 Excel 文件路径：

```bash
python excel_to_ics.py path_to_excel_file.xlsx
```

2. 不使用命令行参数，脚本将使用默认的 Excel 文件路径 `exam_info.xlsx` 并生成 `exam_schedule.ics` 文件：

```bash
python script_name.py
```

### Excel 文件格式

Excel 文件应包含以下列：

- `课程名称`：课程的名称
- `考试地点`：考试的地点
- `考试时间`：考试的时间，格式为 `YYYY-MM-DD (HH:MM-HH:MM)`

### 示例

假设你的 Excel 文件 `exam_info.xlsx` 内容如下：

| 课程名称 | 考试地点 | 考试时间                |
| -------- | -------- | ----------------------- |
| 数学     | 教室101  | 2023-12-20 (09:00-11:00) |
| 英语     | 教室102  | 2023-12-21 (13:00-15:00) |

运行脚本后，将生成一个名为 `exam_schedule.ics` 的 ICS 文件，包含所有考试安排。

### 注意事项

- 确保 Excel 文件中的日期和时间格式正确。
- 脚本默认使用 `Asia/Shanghai` 时区。