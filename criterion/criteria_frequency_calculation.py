import pandas as pd

# 读取 Excel 文件
excel_file = r'E:\github\new-productivity\criterion\criteria_reference_all.xlsx'
sheet_name = '原始指标'
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# 指定要搜索的关键词,只在第A列中搜索





keywords = ['绿色能源','清洁能源消耗','可再生']



# 筛选包含所有关键词的行
filtered_df = df[df.iloc[:, 0].apply(lambda x: any(keyword in str(x) for keyword in keywords))]


# 输出筛选结果
print(filtered_df)
# 输出一下筛选结果的行数
print(filtered_df.shape[0])
