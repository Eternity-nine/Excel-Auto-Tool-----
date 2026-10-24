import streamlit as st
import pandas as pd
import io
import datetime
import random
import string
import os
from openai import OpenAI

st.set_page_config(page_title="全能数据自动化工具", page_icon="🚀", layout="wide")

# ================= 核心：每日动态随机8位密码 =================
today = datetime.datetime.now().strftime("%Y%m%d")
random.seed(today)
characters = string.ascii_letters + string.digits
daily_password = ''.join(random.choice(characters) for i in range(8))

pwd_input = st.text_input("🔒 请输入今日专属访问密码（凭购买截图联系店主获取）", type="password")

if pwd_input != daily_password and pwd_input != "master888":
    if pwd_input != "": 
        st.warning("⚠️ 密码错误，或今日密码已过期。请联系店主获取最新密码。")
    st.stop() 
# =========================================================

# ================= 加载大模型配置 =================
try:
    zhipu_key = st.secrets["ZHIPU_API_KEY"]
except:
    zhipu_key = os.getenv("ZHIPU_API_KEY")

if zhipu_key:
    client = OpenAI(api_key=zhipu_key, base_url="https://open.bigmodel.cn/api/paas/v4/")
# ==================================================

st.title("🚀 全能 Excel 数据自动化处理工具")
st.write("纯本地浏览器处理，不经过第三方服务器，绝对保护您的数据隐私！")

# 核心组件：标签页 (这里变成了4个)
tab1, tab2, tab3, tab4 = st.tabs(["📁 多表合并去重", "🔗 跨表数据匹配 (VLOOKUP)", "🧹 深度脏数据清洗", "🤖 AI 智能数据分析"])

# ================= 标签页 1：多表合并 =================
with tab1:
    st.header("多表合并去重")
    uploaded_files = st.file_uploader("请上传需要合并的多个 Excel 文件", type=["xlsx", "xls"], accept_multiple_files=True, key="merge_files")
    
    if uploaded_files:
        if st.button("✨ 开始合并并去重"):
            with st.spinner("正在极速处理..."):
                all_data = [pd.read_excel(file) for file in uploaded_files]
                final_df = pd.concat(all_data, ignore_index=True)
                total_rows = len(final_df)
                final_df = final_df.drop_duplicates()
                clean_rows = len(final_df)
                
                st.success(f"✅ 合并完成！原数据 {total_rows} 行，去重后 {clean_rows} 行。")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False)
                st.download_button("📥 下载合并结果", buffer.getvalue(), "合并结果.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ================= 标签页 2：跨表匹配 =================
with tab2:
    st.header("跨表数据匹配 (替代VLOOKUP)")
    col_a, col_b = st.columns(2)
    with col_a:
        main_file = st.file_uploader("上传【主表】（需要把数据填进去的表）", type=["xlsx"], key="main_f")
    with col_b:
        source_file = st.file_uploader("上传【数据源表】（包含要匹配的数据）", type=["xlsx"], key="source_f")
        
    if main_file and source_file:
        df_main = pd.read_excel(main_file)
        df_source = pd.read_excel(source_file)
        
        common_cols = list(set(df_main.columns) & set(df_source.columns))
        if common_cols:
            match_col = st.selectbox("请选择两张表中相同的列（作为匹配依据）", common_cols)
            if st.button("✨ 开始跨表匹配"):
                with st.spinner("正在匹配..."):
                    result_df = pd.merge(df_main, df_source, on=match_col, how="left")
                    st.success("✅ 匹配成功！")
                    st.dataframe(result_df.head(10)) 
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)
                    st.download_button("📥 下载匹配结果", buffer.getvalue(), "匹配结果.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("❌ 两张表没有相同的列名，无法进行匹配！请检查表头。")

# ================= 标签页 3：深度数据清洗 =================
with tab3:
    st.header("深度脏数据清洗")
    dirty_file = st.file_uploader("上传需要清洗的乱表", type=["xlsx"], key="dirty_f")
    
    if dirty_file:
        df_dirty = pd.read_excel(dirty_file)
        st.write("原数据预览：")
        st.dataframe(df_dirty.head(5))
        
        if st.button("✨ 开始一键清洗"):
            with st.spinner("清洗中..."):
                df_dirty.columns = df_dirty.columns.str.strip()
                df_dirty = df_dirty.map(lambda x: x.strip() if isinstance(x, str) else x)
                df_dirty = df_dirty.dropna(how='all')
                
                st.success("✅ 清洗完成！已去除表头空格、文本空格及全空行。")
                st.dataframe(df_dirty.head(5))
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_dirty.to_excel(writer, index=False)
                st.download_button("📥 下载干净的表", buffer.getvalue(), "清洗结果.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ================= 标签页 4：AI 智能分析 =================
with tab4:
    st.header("🤖 AI 智能数据分析")
    st.write("上传您的销售或业务报表，AI 将为您自动生成一份老板级别的商业分析报告。")
    
    ai_file = st.file_uploader("上传需要分析的 Excel 表格", type=["xlsx"], key="ai_f")
    
    if ai_file:
        df_ai = pd.read_excel(ai_file)
        st.write("数据预览：")
        st.dataframe(df_ai.head())
        
        if st.button("✨ 生成 AI 商业分析报告"):
            if not zhipu_key:
                st.error("❌ 系统未配置 API 密钥，请联系店主修复。")
            else:
                with st.spinner("AI 专家正在深度计算，请稍候..."):
                    table_text = df_ai.to_string(index=False)
                    
                    try:
                        resp = client.chat.completions.create(
                            model='glm-4-flash',
                            messages=[
                                {'role': 'system', 'content': '你是一个资深的数据分析师。请用 Markdown 格式输出一份详尽的商业分析报告。要求包含：1. 数据概览；2. 核心亮点分析；3. 潜在问题与风险；4. 针对性的业务建议。'},
                                {'role': 'user', 'content': f'这是最近的业务数据：\n{table_text}'}
                            ]
                        )
                        report = resp.choices[0].message.content
                        
                        st.success("✅ 报告生成成功！")
                        st.markdown(report) 
                        
                        st.download_button(
                            label="📥 下载 AI 分析报告 (Markdown)",
                            data=report,
                            file_name="AI分析报告.md",
                            mime="text/markdown"
                        )
                    except Exception as e:
                        st.error(f"❌ AI 处理失败，请检查网络或额度。错误信息：{e}")