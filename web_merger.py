import streamlit as st
import pandas as pd
import io
import datetime

st.set_page_config(page_title="全能数据自动化工具", page_icon="🚀", layout="wide")

# ================= 核心：动态防盗版密码机制 =================
today = datetime.datetime.now().strftime("%m%d")
daily_password = f"Excel{today}" # 比如今天是10月24日，密码就是 Excel1024

pwd_input = st.text_input("🔒 请输入今日专属访问密码（凭购买截图联系店主获取）", type="password")

# 校验逻辑：如果输入的密码不对，并且也不是万能密码，直接掐断程序！
if pwd_input != daily_password and pwd_input != "master888":
    if pwd_input != "": # 如果用户输入了密码但是错的，给出警告
        st.warning("⚠️ 密码错误，或今日密码已过期。请联系店主获取最新密码。")
    st.stop() # 核心防线：没有这行，防盗就失败了。必须保留并保持缩进！
# =========================================================

# 以下所有代码，只有密码输入正确后，才会显示
st.title("🚀 全能 Excel 数据自动化处理工具")
st.write("纯本地浏览器处理，不经过第三方服务器，绝对保护您的数据隐私！")

# 核心组件：标签页
tab1, tab2, tab3 = st.tabs(["📁 多表合并去重", "🔗 跨表数据匹配 (VLOOKUP)", "🧹 深度脏数据清洗"])

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
                # 注意这里的括号和引号已经修补完整
                st.download_button("📥 下载干净的表", buffer.getvalue(), "清洗结果.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")