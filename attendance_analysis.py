import pandas as pd
import re

def parse_attendance():

    # Load the Excel file
    xls = pd.ExcelFile("数据源\考勤.xlsx")

    # Initialize a dictionary to store results
    attendance_dict = {}

    # 初始化存储结果的字典
    employment_dates = {}

    # Process each sheet
    for sheet_name in xls.sheet_names:
        # Load the sheet into a dataframe
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # 清理列名中的不需要字符，例如"_x000D_"
        df.columns = df.columns.str.replace(r'_x000D_|\n', '', regex=True)
        
        # Ensure column names are unique
        df.columns = pd.Series(df.columns).apply(lambda x: x if x not in df.columns[:df.columns.get_loc(x)] else f"{x}_{df.columns.get_loc(x)}")
        
        # Identify key columns
        name_col = None
        entry_date_col = None
        exit_date_col = None
        
        for col in df.columns:
            if "姓名" in col:
                name_col = col
            elif "入职" in col:
                entry_date_col = col
            elif "离职" in col:
                exit_date_col = col
        
        if not all([name_col, entry_date_col, exit_date_col]):
            # Skip if any of the key columns is not found
            raise ValueError(f"在工作表 {sheet_name} 中未找到必要的列（姓名、入职时间）。")
        

        # 遍历所有行，分析入职和离职时间
        for index, row in df.iterrows():
            name = row[name_col]
            entry_date = pd.to_datetime(row[entry_date_col], errors='coerce')
            exit_date = pd.to_datetime(row[exit_date_col], errors='coerce') if pd.notna(row[exit_date_col]) else pd.Timestamp.max

            if pd.isna(entry_date):
                raise ValueError(f"在工作表 {sheet_name} 的第 {index + 1} 行，姓名为 {name} 的员工入职时间为空。")
            
            # 更新字典中的入职和离职日期范围
            if name not in employment_dates:
                employment_dates[name] = {'入职时间': entry_date, '离职时间': exit_date}
            else:
                employment_dates[name]['入职时间'] = min(employment_dates[name]['入职时间'], entry_date)
                employment_dates[name]['离职时间'] = min(employment_dates[name]['离职时间'], exit_date)


        # 将结果转换为字典
        employment_dates_dict = pd.DataFrame.from_dict(employment_dates, orient='index').to_dict('index')

        # Convert '入职时间' and '离职时间' to datetime
        df[entry_date_col] = pd.to_datetime(df[entry_date_col], errors='coerce')
        df[exit_date_col] = pd.to_datetime(df[exit_date_col], errors='coerce')
        
        # Process date columns
        for col in df.columns[3:]:
            # Extract the day number from the column name
            date_match = re.search(r'(\d+)', col)
            if date_match:
                day = date_match.group(1)
                current_date = pd.to_datetime(f"{sheet_name[:4]}-{sheet_name[4:]}-{day.zfill(2)}", errors='coerce')
                
                # Filter rows that are within the valid employment date range
                valid_rows = df[
                    (df[entry_date_col].isna() | (current_date >= df[entry_date_col])) &
                    (df[exit_date_col].isna() | (current_date <= df[exit_date_col]))
                ]
                
                # Identify absences where the value is neither '正常' nor '加班'
                for index, row in valid_rows.iterrows():
                    value = row[col]
                    if pd.notna(value) and \
                        ("休息" not in value and "正常" not in value and "加班" not in value and "漏签" not in value):
                        name = row[name_col]
                        if name not in attendance_dict:
                            attendance_dict[name] = []
                        attendance_dict[name].append((current_date.strftime('%Y-%m-%d'), value))

    return attendance_dict, employment_dates_dict


if __name__ == '__main__':
    att, emp = parse_attendance()