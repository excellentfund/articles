import fnmatch
import os
import pandas as pd
import re


# 定义修约间隔函数
def get_rounding_interval(value):
    if 0 <= value < 2:
        return 0.05
    elif 2 <= value < 5:
        return 0.1
    elif 5 <= value < 10:
        return 0.25
    elif 10 <= value < 20:
        return 0.5
    elif 20 <= value < 50:
        return 1
    elif 50 <= value < 200:
        return 2.5
    elif 200 <= value < 300:
        return 5
    elif 300 <= value < 500:
        return 10
    elif value >= 500:
        return 10
    else:
        raise ValueError(f"Value {value} is valid")


# 读取Excel文件中的B列数据
excel_file_path = r'C:\Users\workstation\Desktop\doctest\hkoptions.xlsx'
df = pd.read_excel(excel_file_path)

# 将B列数据转换为字符串，并在不足5位的前面补0
column_b_data = df.iloc[:, 1].astype(str).apply(lambda x: x.zfill(5))

# 初始化结果列表
results = []

# 遍历B列数据
for value in column_b_data:
    # 构建文件搜索路径，查找包含股票代码的文件
    search_dir = r'C:\Users\workstation\Desktop\doctest'
    file_pattern = f'*{value}*.txt'

    # 查找匹配的txt文件
    for filename in os.listdir(search_dir):
        if fnmatch.fnmatch(filename, file_pattern):
            file_path = os.path.join(search_dir, filename)

            # 读取txt文件内容。注：win环境下txt文件编码为ASNI（cp1252）
            with open(file_path, 'r', encoding='cp1252') as file:
                for line in file:
                    # 使用正则表达式匹配行内容
                    match = re.match(r'^(\d{4}/\d{2}/\d{2}),(.+),(.+),(.+),(.+),(.+),(.+)$', line.strip())
                    if match:
                        groups = match.groups()
                        date = groups[0]
                        if date == '2024/10/03':
                            close_data = groups[4]
                            close_data_float = float(close_data)
                            # 计算修约间隔
                            rounding_interval = get_rounding_interval(close_data_float)
                            # 四舍五入为整数，再乘以修约间隔
                            rounded_value = round(close_data_float / rounding_interval) * rounding_interval
                            # 添加结果
                            results.append([value, date, close_data, rounded_value])
                            # 不再找后续行
                            break
            # 如果找到符合的文件，不再找后续
            break

# 创建DataFrame并保存到新的Excel文件
output_df = pd.DataFrame(results, columns=['股票代码', '日期', '收盘', '平值期权行权价'])
output_df.to_excel(r'C:\Users\workstation\Desktop\doctest\output.xlsx', index=False)

print("处理完成，结果已保存到output.xlsx")
