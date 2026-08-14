import streamlit as st
import pandas as pd
import io
import itertools
import re
from rapidfuzz import fuzz

st.set_page_config(page_title="设备表格核对工具（多条件+地址结构保护）", layout="wide")
st.title("🔍 设备表格核对工具（多条件+地址结构保护）")
st.write("上传你的设备表和客户设备表，支持多个匹配条件；地址列可选择“模糊+地址保护”，确保同级别比较，避免跨级误匹配。")

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

    st.subheader("设置匹配条件")
    st.info("""
    - **精确匹配**：单元格内容完全一致（忽略大小写和首尾空格）。支持我的表格中一个单元格包含多个值（如“1,2,3”），会自动拆分。
    - **模糊匹配**：适合普通文本，使用综合相似度算法（WRatio）。
    - **模糊+地址保护**：适合地址列，在模糊匹配前先检查省、市、区、街道、路等层级是否一致或兼容，防止跨级误配。
    - 未启用的条件请保持“— 不启用 —”。
    """)

    # 模糊匹配阈值滑块（仅用于模糊匹配和模糊+地址保护）
    fuzzy_threshold = st.slider("模糊匹配相似度阈值（%）", min_value=50, max_value=100, value=70, step=1,
                                help="只有相似度超过该值才视为匹配，值越高越严格，越低越宽松。建议65~85之间。")

    # 数据清洗函数
    def clean_str(s):
        if pd.isna(s):
            return ''
        return str(s).strip().upper()

    # 拆分多值函数（用于精确匹配）
    def split_if_multi(value):
        if pd.isna(value):
            return ['']
        if re.search(r'[，,;；、\s|/]+', str(value)):
            parts = re.split(r'[，,;；、\s|/]+', str(value))
            parts = [p.strip() for p in parts if p.strip()]
            return parts if parts else [str(value)]
        else:
            return [str(value)]

    # ---------- 地址结构保护相关函数 ----------
    def extract_level_keys(text):
        """提取地址中的省、市、区、街道、路等层级的关键词主体"""
        text = str(text)
        patterns = {
            'province': r'([\u4e00-\u9fa5]{2,8}(?:省|自治区|特别行政区))',
            'city': r'([\u4e00-\u9fa5]{2,8}(?:市|自治州|地区|盟))',
            'district': r'([\u4e00-\u9fa5]{2,8}(?:区|县|旗|新区))',
            'street': r'([\u4e00-\u9fa5]{2,8}(?:街道|镇|乡))',
            'road': r'([\u4e00-\u9fa5]{2,12}(?:路|街|大道|巷|道))'
        }
        suffixes = ['特别行政区', '自治区', '自治州', '地区', '盟', '新区', '省', '市', '区', '县', '旗', '街道', '镇', '乡', '路', '街', '大道', '巷', '道']
        levels = {}
        for level, pattern in patterns.items():
            matches = re.findall(pattern, text)
            cleaned = []
            for m in matches:
                # 去掉后缀，保留主体
                for suf in suffixes:
                    if m.endswith(suf):
                        cleaned.append(m[:-len(suf)])
                        break
                else:
                    cleaned.append(m)
            levels[level] = cleaned
        return levels

    def address_level_check(addr1, addr2):
        """检查两个地址的层级是否冲突：同层级若都存在则必须有交集"""
        levels1 = extract_level_keys(addr1)
        levels2 = extract_level_keys(addr2)
        for level in ['province', 'city', 'district', 'street', 'road']:
            list1 = set(levels1.get(level, []))
            list2 = set(levels2.get(level, []))
            if list1 and list2:
                # 同层级两边都有，但无交集则冲突
                if not list1 & list2:
                    return False
        return True
    # ----------------------------------------

    # 准备列选项
    none_option = "— 不启用 —"
    my_cols = [none_option] + list(my_df.columns)
    cust_cols = [none_option] + list(customer_df.columns)

    # 创建4组条件，每组增加匹配方式选择
    conditions = []
    for i in range(4):
        col_a, col_b, col_c = st.columns([1, 1, 0.9])
        with col_a:
            my_col = st.selectbox(f"条件{i+1}：我的表格列", my_cols, key=f"my_col_{i}")
        with col_b:
            cust_col = st.selectbox(f"条件{i+1}：客户表格列", cust_cols, key=f"cust_col_{i}")
        with col_c:
            match_type = st.selectbox(f"匹配方式", ["精确", "模糊", "模糊+地址保护"], key=f"match_type_{i}")
        conditions.append((my_col, cust_col, match_type))

    # 检查是否至少启用了一个条件
    active_conditions = [(my, cust, mtype) for my, cust, mtype in conditions if my != none_option and cust != none_option]

    if st.button("🚀 开始匹配"):
        if not active_conditions:
            st.error("请至少设置一个匹配条件（两个表格的列都要选择）。")
        else:
            # 预处理我的表格数据
            my_rows = []
            for _, row in my_df.iterrows():
                row_data = []
                for my_col, _, mtype in active_conditions:
                    val = row[my_col]
                    if mtype == "精确":
                        vals = [clean_str(x) for x in split_if_multi(val)]
                        row_data.append(vals)
                    else:
                        row_data.append(clean_str(val))
                my_rows.append(row_data)

            # 逐行匹配客户表格
            matched_indices = []
            for cust_idx, cust_row in customer_df.iterrows():
                is_match = False
                for my_row in my_rows:
                    all_pass = True
                    for cond_idx, (my_col, cust_col, mtype) in enumerate(active_conditions):
                        cust_val = clean_str(cust_row[cust_col])
                        if mtype == "精确":
                            if cust_val not in my_row[cond_idx]:
                                all_pass = False
                                break
                        else:
                            my_val = my_row[cond_idx]
                            # 模糊+地址保护：先检查地址结构
                            if mtype == "模糊+地址保护":
                                if not address_level_check(my_val, cust_val):
                                    all_pass = False
                                    break
                            # 计算相似度（WRatio）
                            if not my_val or not cust_val:
                                score = 0
                            else:
                                score = fuzz.WRatio(my_val, cust_val)
                            if score < fuzzy_threshold:
                                all_pass = False
                                break
                    if all_pass:
                        is_match = True
                        break
                if is_match:
                    matched_indices.append(cust_idx)

            matched_df = customer_df.loc[matched_indices].copy()

            st.success(f"✅ 匹配完成！客户表格中共有 {len(matched_df)} 条记录属于你的设备。")

            if len(matched_df) > 0:
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
                st.warning("没有找到匹配的设备，请调整匹配条件或阈值后重试。")
