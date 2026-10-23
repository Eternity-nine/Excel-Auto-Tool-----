import pandas as pd
import os

# 假设客户给你发了两个表
file_a = '客户原表/员工信息表.xlsx'  # 主表（需要把数据填进去的表）
file_b = '客户原表/工资表.xlsx'      # 数据源表（包含需要匹配的数据）

output_folder = '处理结果'

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

try:
    print(f"🔍 正在读取主表：{file_a}")
    df_a = pd.read_excel(file_a)
    print(f"🔍 正在读取数据源表：{file_b}")
    df_b = pd.read_excel(file_b)
    
    # 核心动作：跨表匹配！
    # on='姓名' 意思是：用两表都有的“姓名”列作为匹配依据
    # how='left' 意思是：保留主表的所有人，只有匹配不到的才留空（就像 VLOOKUP 一样）
    print("🔗 正在执行跨表匹配...")
    final_df = pd.merge(df_a, df_b, on='姓名', how='left')
    
    # 保存
    save_path = os.path.join(output_folder, 'VLOOKUP匹配结果.xlsx')
    final_df.to_excel(save_path, index=False)
    
    print("-" * 30)
    print(f"✅ 匹配成功！")
    print(f"结果已保存到：{save_path}")

except KeyError as e:
    print(f"❌ 发生错误：找不到列名 {e}。请确保两个表里都有这一列，并且名字一模一样。")
except Exception as e:
    print(f"❌ 发生错误：{e}")