import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="设备表格核对工具（多条件版）", layout="wide")
st.title("🔍 设备表格核对工具（多条件版）")
st.write("上传你的设备表和客户设备表，可设置多个匹配条件，精确找出属于你的设备。")

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
        my_df = load_file(my_file).reset_index(drop=True)
        customer_df = load_file(customer_file).reset_index(drop=True)
    except Exception as e:
        st.error(f"读取文件出错：{e}")
        st.stop()

    st.subheader("设置匹配条件（可同时匹配多列）")
    st.info("请为每个匹配条件选择“我的表格”和“客户表格”中对应的列。\n"
            "例如：条件1：我的表格的“设备号” ↔ 客户表格的“设备号”；条件2：我的表格的“城市” ↔ 客户表格的“城市”。\n"
            "未启用的条件请保持“— 不启用 —”。")

    # 数据清洗函数：去除空格、统一大写、填充空值
    def clean_series(s):
        return s.fillna('').astype(str).str.strip().str.upper()

    # 准备列选项
    none_option = "— 不启用 —"
    my_cols = [none_option] + list(my_df.columns)
    cust_cols = [none_option] + list(customer_df.columns)

    # 创建4组条件
    conditions = []
    for i in range(4):
        col_a, col_b, col_c = st.columns([1, 1, 0.1])
        with col_a:
            my_col = st.selectbox(f"条件{i+1}：我的表格列", my_cols, key=f"my_col_{i}")
        with col_b:
            cust_col = st.selectbox(f"条件{i+1}：客户表格列", cust_cols, key=f"cust_col_{i}")
        conditions.append((my_col, cust_col))

    # 检查是否至少启用了一个条件
    active_conditions = [(my, cust) for my, cust in conditions if my != none_option and cust != none_option]

    if st.button("🚀 开始匹配"):
        if not active_conditions:
            st.error("请至少设置一个匹配条件（两个表格的列都要选择）。")
        else:
            # 确保所有列都存在（已通过选项保证，但再确认一次）
            for my_col, cust_col in active_conditions:
                if my_col not in my_df.columns or cust_col not in customer_df.columns:
                    st.error(f"列名不存在：我的表格中的 {my_col} 或客户表格中的 {cust_col}")
                    st.stop()

            # 生成匹配键集合：将“我的表格”中指定列的组合转为元组集合
            my_key_cols = [my_col for my_col, _ in active_conditions]
            my_keys = my_df[my_key_cols].apply(lambda row: tuple(clean_series(row)), axis=1)
            my_key_set = set(my_keys)

            # 生成客户表格的对应键，并判断是否在集合中
            cust_key_cols = [cust_col for _, cust_col in active_conditions]
            cust_keys = customer_df[cust_key_cols].apply(lambda row: tuple(clean_series(row)), axis=1)
            mask = cust_keys.isin(my_key_set)
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
                st.warning("没有找到匹配的设备，请检查匹配条件或数据格式。")
