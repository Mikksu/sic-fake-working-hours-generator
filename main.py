import pandas as pd
import matplotlib.pyplot as plt
from attendance_analysis import parse_attendance


#! 注意，先尝试打开输出报表，如果打不开则不做下面的工作，以节省时间
#writer = pd.ExcelWriter("output.xlsx")
#writer.handles = None
#writer.close()

print('分析考勤记录...')
attendance_dict, employment_dates_dict = parse_attendance()

print('分配工时 ....')

import pandas as pd
from random import choice

# 加载 Excel 文件
file_path = '数据源/项目工时人员分配.xlsx'
xls = pd.ExcelFile(file_path)

# Step 1: 统计项目名称
projects = xls.sheet_names

# 每人每月所在项目
person_month_project = {}

# 所有项目的月份列表
month_list_all_project = []

# 每个项目月份范围
project_month_range = {}

# Step 2: 统计所有项目中的人员清单
names_list = set()
for project in projects:
    df = pd.read_excel(xls, sheet_name=project)

    # 获取时间列表
    all_times = df.iloc[0]

    # 创建项目日期范围字典
    if project in project_month_range:
        raise ValueError
    project_month_range[project] = {'Begin': all_times.index[0], 'End': pd.offsets.MonthEnd().rollforward(all_times.index[-1])}

    for i in range(len(all_times.index)):

        # 第i列月份
        date_str = all_times.index[i]
        try:
            month = pd.to_datetime(date_str, format='%Y-%m', errors='coerce')
        except ValueError:
            raise

        # 所有项目月份列表
        if month not in month_list_all_project:
            month_list_all_project.append(month)

        # 当前月份人员
        names_in_month = df.iloc[0:, i].dropna()
        
        # 创建[人员][月份]下的项目列表
        for person in names_in_month:
            if person not in person_month_project:
                person_month_project[person] = {}
            if month not in person_month_project[person]:
                person_month_project[person][month] = []
            if project not in person_month_project[person][month]:
                person_month_project[person][month].append(project)


    # 获取姓名列表
    names_curr_project = df.iloc[0:, :].values.flatten().tolist()

    # 过滤掉None或NaN值
    names_curr_project = [name for name in names_curr_project if pd.notna(name)]

    # 去重
    names_list.update(names_curr_project)

names_list = list(names_list)

print(f'共解析 {len(project_month_range)} 个项目。')


# Step 3: 按项目分类统计并分配工时
project_hours = {project: {} for project in projects}

# today
today = pd.Timestamp.now()

for month in month_list_all_project:
    ts = pd.Timestamp(month)

    # 获取该月的所有日期
    month_start = ts.replace(day=1)
    month_end = pd.offsets.MonthEnd().rollforward(month_start)

    # 使用date_range生成该月份的所有日期
    dates = pd.date_range(start=month_start, end=month_end)

    # 提取当月天数列表
    days_list = dates.tolist()

    for curr_day in days_list:
        print(f'正在分析{curr_day.strftime('%Y-%m-%d')}\r', end='')

        # 如果今天是周六或周日，跳过
        if curr_day.dayofweek == 0 or curr_day.dayofweek == 6:
            continue

        # 如果日期超过今天日期，则不分配工时
        if curr_day >= today:
            continue

        #遍历总项目清单中的每个人
        for person in person_month_project:

            # 如果这个人当天没有出勤，则不要统计工时
            if person in attendance_dict:
                att_list_person = attendance_dict[person]
                ret = [item for item in att_list_person if item[0] == curr_day.strftime('%Y-%m-%d')]
                if any(ret):
                    continue

            # 如果这个人当天未入职或已离职，则不要统计工时
            if person in employment_dates_dict:
                entry_date = employment_dates_dict[person]['入职时间']
                exit_date = employment_dates_dict[person]['离职时间']
                if not (entry_date <= curr_day <= exit_date):
                    continue

            # 如果这个人在当前月份中没有项目，则跳过
            month_belongs_to = curr_day.replace(day=1)
            if month_belongs_to not in person_month_project[person]:
                #print(f'日期{month_belongs_to}中{person}没有项目。')
                continue

            # 获取这个人在这个月所在的所有项目清单
            projs = person_month_project[person][month_belongs_to]

            # 随机挑选一个项目，作为当天工时统计项
            chosen_proj = choice(projs)

            # 工时归属到当前项目下的这个人名下
            if chosen_proj not in project_hours:
                project_hours[chosen_proj] = {}
            if person not in project_hours[chosen_proj]:
                project_hours[chosen_proj][person] = []

            project_hours[chosen_proj][person].append(curr_day)



with pd.ExcelWriter("output.xlsx") as writer:
    # 按项目名称创建excel
    for curr_proj in project_hours:
        
        print(f'创建项目 {curr_proj} 的工时报表...', end='\n')
        rd_hours_summary = {}

        # 当前项目日期范围
        proj_start = project_month_range[curr_proj]['Begin']
        proj_end = project_month_range[curr_proj]['End']

        # 按天生成当前项目日期列表
        proj_days = pd.date_range(start=proj_start, end=proj_end)

        for person in project_hours[curr_proj]:
            rd_hours_summary[person] = {}
            for day in proj_days:
                if day in project_hours[curr_proj][person]:
                    rd_hours_summary[person][day.strftime('%Y-%m-%d')] = 8
                else:
                    rd_hours_summary[person][day.strftime('%Y-%m-%d')] = None


        df = pd.DataFrame.from_dict(rd_hours_summary)
        df.to_excel(writer, sheet_name=curr_proj)  



print('完成工时分配')