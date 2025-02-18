# Exam Schedule to ICS Converter

[**exam_to_ics2**](#exam_to_ics2)是最新版本，无需用户从教学信息服务网导出考试安排，通过`pysjtu`库抓取考试安排，并数据源转换为 ICS 文件，以便导入到日历应用中。

[**excel_to_ics**](#excel_to_ics)用于将 SJTU 教学信息服务网导出的考试安排从 Excel 文件转换为 ICS 文件，以便导入到日历应用中。

## `exam_to_ics2`使用方法 {#exam_to_ics2}

### 依赖安装

在运行脚本之前，请确保已安装以下依赖项：

```bash
pip install icalender pytz pysjtu
```

### 运行脚本

你可以通过以下两种方式运行脚本：

1. 使用命令行参数指定Jaccount账号、密码和学年、学期

    ```bash
    python exam_to_ics2.py username password year term
    ```
    例如：

    ```bash
    python exam_to_ics2.py MyJaccount sjtu1896 2024 1
    ```

2. 不使用命令行参数，脚本将提示输入Jaccount账号、密码和学年、学期

    ```bash
    python exam_to_ics2.py
    ```

3. 运行脚本后，将在运行目录下创建一个名为`data`的文件夹，并在其中生成一个名为 `{学年}-{学期}考试安排.ics` 的 ICS 文件，包含所有考试安排。

### 数据格式

脚本假设考试安排信息存储在一个列表中，每个考试信息是一个字典，包含以下字段：

- `course_name`：课程的名称
- `course_id`：课程的编号
- `location`：考试的地点
- `date`：考试的日期，格式为 `[datetime.data]`
- `time`：考试的时间，格式为 `[datetime.time(开始时间), datetime.time(结束时间)]`

### 注意事项

- 确保日期和时间格式正确。
- 脚本默认使用 `Asia/Shanghai` 时区。


## `excel_to_ics`使用方法 {#excel_to_ics}

### 依赖安装

在运行脚本之前，请确保已安装以下依赖项：

```bash
pip install pandas icalendar pytz
```

### 运行脚本

你可以通过以下两种方式运行脚本：

1. 使用命令行参数指定 Excel 文件路径：

    ```bash
    python excel_to_ics.py path_to_excel_file.xlsx
    ```
    例如：

    ```bash
    python excel_to_ics.py ./data/2024-2025-1.xlsx
    ```

2. 不使用命令行参数，脚本将提示输入文件路径（支持 Excel 或 txt 文件），并生成相应的 ICS 文件：

    ```bash
    python excel_to_ics.py
    ```

    在运行脚本后，输入文件路径，例如：

    ```plaintext
    请输入文件路径（支持Excel或txt文件）：./data/2024-2025-1.xlsx
    ```

3. 运行脚本后，将在存放 Excel 文件或 txt 文件的目录生成一个名为同名的 ICS 文件，包含所有考试安排。

### 文件格式

Excel 文件或 txt 文件应包含以下列：

- `课程代码`: 课程的代码
- `课程名称`：课程的名称
- `考试地点`：考试的地点
- `考试时间`：考试的时间，格式为 `YYYY-MM-DD (HH:MM-HH:MM)`

### 示例
Excel 文件 或 txt 文件可以从 [上海交通大学教学信息服务网 > 信息查询 > 考试信息查询](https://i.sjtu.edu.cn/kwgl/kscx_cxXsksxxIndex.html?gnmkdm=N358105&layout=default) 导出。

假设你的 Excel 文件 `exam_info.xlsx` 或 txt 文件 `exam_info.txt` 内容如下：

| 课程代码 | 课程名称                 | 考试地点 | 考试时间                |
| -------- | ------------------------ | -------- | ----------------------- |
| CS2501   | 离散数学                 | 上院114  | 2025-01-01 (11:40-15:40) |
| PHY1252  | 大学物理                 | 中院514  | 2025-01-02 (11:40-15:40) |
| MATH1207 | 概率统计                 | 上院315  | 2025-01-03 (11:40-15:40) |

### 注意事项

- 确保 Excel 文件或 txt 文件中的日期和时间格式正确。
- 脚本默认使用 `Asia/Shanghai` 时区。
