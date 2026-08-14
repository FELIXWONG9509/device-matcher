import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="设备表格核对工具", layout="wide")
st.title("🔍 设备表格核对工具")
st.write("上传你的设备表和客户设备表，自动找出属于你的设备。")

# 上传文件
col1, col2 = st.columns(2)
with col1:
    my_file = st.file_uploader("📤 上传你的设备表（Excel或CSV）", type=["xlsx", "xls", "csv"], key="my")
with col2:
    customer_file = st.file_uploader("📤 上传客户的设备表（Excel或CSV）", type=["xlsx", "xls", "csv"], key="customer")

if my_file and customer_file:
    # 读取文件函数
    def load_file(file):
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        elif file.name.endswith('.xls'):
            return pd.read_excel(file, engine='xlrd')
        else:
            return pd.read_excel(file, engine='openpyxl')

    try:
        my_df = load_file(my_file)
        customer_df = load_file(customer_file)
    except Exception as e:
        st.error(f"读取文件出错：{e}")
        st.stop()

    st.subheader("选择用于匹配的列")
    st.info("请分别选择两个表格中代表“设备号”的列（列名可以不同）。")
    col3, col4 = st.columns(2)
    with col3:
        my_col = st.selectbox("你的表格中的设备号列", my_df.columns, key="my_col")
    with col4:
        customer_col = st.selectbox("客户表格中的设备号列", customer_df.columns, key="customer_col")

    # 数据清洗函数：去除空格、统一大写、处理空值
    def clean_series(s):
        return s.fillna('').astype(str).str.strip().str.upper()

    if st.button("🚀 开始匹配"):
        # 提取你的设备号（去重）
        my_ids = clean_series(my_df[my_col]).drop_duplicates()
        # 客户表设备号
        customer_ids = clean_series(customer_df[customer_col])
        # 匹配
        mask = customer_ids.isin(my_ids)
        matched_df = customer_df[mask].copy()

        st.success(f"✅ 匹配完成！客户表格中共有 {len(matched_df)} 条记录属于你的设备。")

        if len(matched_df) > 0:
            # 生成 Excel 文件供下载
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                matched_df.to_excel(writer, index=False, sheet_name='匹配结果')
            output.seek(0)

            st.download_button(
                label="📥 下载匹配结果 (Excel)",
                data=output,
                file_name="匹配结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("没有找到匹配的设备，请检查列名或数据格式。")
