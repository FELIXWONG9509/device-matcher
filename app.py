import streamlit as st
import pandas as pd
import io
import itertools
import re

st.set_page_config(page_title="设备表格核对工具（多条件+多值拆分）", layout="wide")
st.title("🔍 设备表格核对工具（多条件+多值拆分）")
st.write("上传你的设备表和客户设备表，支持多个匹配条件，并且我的表格中某列包含多个值时，会自动拆分匹配。")

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
            "未启用的条件请保持“— 不启用 —”。\n"
            "💡 如果你的表格某列（如“设备号”）的一个单元格里包含多个值（例如“1,2,3”），程序会自动拆分并匹配客户表格中的单个值。")

    # 数据清洗函数：去除空格、统一大写、处理空值
    def clean_str(s):
        if pd.isna(s):
            return ''
        return str(s).strip().upper()

    # 判断是否包含分隔符，并拆分
    def split_if_multi(value):
        if pd.isna(value):
            return ['']
        # 常见分隔符：逗号（中英文）、分号（中英文）、顿号、空格、竖线、斜杠
        if re.search(r'[，,;；、\s|/]+', str(value)):
            parts = re.split(r'[，,;；、\s|/]+', str(value))
            parts = [p.strip() for p in parts if p.strip()]
            return parts if parts else [str(value)]
        else:
            return [str(value)]

    # 准备列选项
    none_option = "— 不启用 —"
    my_cols = [none_option] + list(my_df.columns)
    cust_cols = [none_option] + list(customer_df.columns)

    # 创建4组条件
    conditions = []
    for i in range(4):
        col_a, col_b = st.columns(2)
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
            # 构建“我的表格”的匹配键集合（按行处理，支持多值拆分）
            my_key_cols = [my_col for my_col, _ in active_conditions]
            my_key_set = set()

            # 遍历我的表格的每一行
            for _, row in my_df[my_key_cols].iterrows():
                # 对当前行的每个单元格进行拆分并清洗，得到列表
                split_lists = []
                for val in row:
                    vals = split_if_multi(val)          # 拆分多值
                    clean_vals = [clean_str(v) for v in vals]  # 清洗每个拆分值
                    split_lists.append(clean_vals)
                # 当前行的笛卡尔积（如果所有单元格都是单值，则只有一个组合）
                for combo in itertools.product(*split_lists):
                    my_key_set.add(tuple(combo))

            # 构建客户表格的匹配键（不拆分，只清洗）
            cust_key_cols = [cust_col for _, cust_col in active_conditions]
            cust_keys = customer_df[cust_key_cols].apply(
                lambda row: tuple(clean_str(x) for x in row), axis=1
            )

            # 匹配
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
