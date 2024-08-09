import pandas as pd
import re

def parse_dep_tech_overtime():
    # Load excel file
    xls = pd.ExcelFile("数据源\\2022-2023技术部周末加班工时.xlsx")
    sheet_name = "跑片记录 (2022-2023)"

    #Initialize dictionary
    dict_overtime_proj_person_date = {}

    name_list = set()

    df = pd.read_excel(xls, sheet_name)

    # Ensure column names are unique
    df.columns = pd.Series(df.columns).apply(lambda x: x if x not in df.columns[:df.columns.get_loc(x)] else f"{x}_{df.columns.get_loc(x)}")

    project_col = None
    person_col = None
    date_col = None

    for col in df.columns:
        if "备注" in col:
            person_col = col
        elif "使用日期" in col:
            date_col = col
        elif "项目号" in col:
            project_col = col

    if not all([project_col, person_col, date_col]):
        raise ValueError(f"工作表 [{sheet_name}] 未找到必要的列。")
    

    for index, row in df.iterrows():
        person = row[person_col]
        project = row[project_col]
        date = row[date_col]

        if pd.isna(person) or pd.isna(project) or pd.isna(date):
            raise ValueError("存在空数据。")

        if person not in dict_overtime_proj_person_date:
            dict_overtime_proj_person_date[person] = []
            name_list.update([person])
        
        day = pd.Timestamp(year=date.year, month=date.month, day=date.day)
        dict_overtime_proj_person_date[person].append((day, project))
    
    return dict_overtime_proj_person_date, name_list
        

if __name__ == '__main__':
    dict = parse_dep_tech_overtime()