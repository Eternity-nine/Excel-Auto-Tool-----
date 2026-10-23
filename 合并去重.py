import pandas as pd
import os

# 定义文件夹路径
input_folder = '客户原表'
output_folder = '处理结果'

# 如果输出文件夹不存在，就自动创建
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取所有 xlsx 结尾的文件，并排除以 "~$" 开头的隐藏临时文件（客户没关表时的坑）
files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') and not f.startswith('~$')]

if len(files) == 0:
    print("⚠️ 报错：客户原表文件夹里没有找到 Excel 文件！")
else:
    print(f"🔍 发现 {len(files)} 个表格，开始处理...")
    
    all_data = []
    # 循环读取每个表
    for file in files:
        file_path = os.path.join(input_folder, file)
        df = pd.read_excel(file_path)
        all_data.append(df)
        print(f"  - 已读取: {file}")
    
    # 核心动作1：把所有表像搭积木一样拼起来
    final_df = pd.concat(all_data, ignore_index=True)
    total_rows = len(final_df)
    
    # 核心动作2：一键去重
    final_df = final_df.drop_duplicates()
    clean_rows = len(final_df)
    
    # 保存结果
    save_path = os.path.join(output_folder, '最终合并结果.xlsx')
    final_df.to_excel(save_path, index=False)
    
    print("-" * 30)
    print(f"✅ 处理完成！")
    print(f"合并前行数: {total_rows} 行")
    print(f"去重后行数: {clean_rows} 行（帮客户删掉了 {total_rows - clean_rows} 条重复数据）")
    print(f"结果已保存到：{save_path}")