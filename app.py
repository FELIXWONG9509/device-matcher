import streamlit as st
import pandas as pd
import io
import itertools
import re
from rapidfuzz import fuzz

st.set_page_config(page_title="设备表格核对工具（多条件+地址保护+自动忽略括号）", layout="wide")
st.title("🔍 设备表格核对工具（多条件+地址保护+自动忽略括号）")
st.write("上传你的设备表和客户设备表，支持多个匹配条件；地址列自动进行同级别比较，并自动忽略括号内容。")

# 上传文件
col1, col2 = st.columns(2)
with col1:
    my_file = st.file_uploader("📤 上传你的设备表（Excel或CSV）", type=["xlsx", "xls", "csv"], key="my")
with col2:
    customer_file = st.file_uploader("📤 上传客户的设备表（Excel或CSV）", type=["xlsx", "xls", "csv"], key="customer")

if my_file and customer_file:
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
    - **模糊匹配**：适合普通文本，使用多种相似度算法取最大值。**自动忽略括号及其中的内容**（例如“（华为超充技术支持）”）。
    - **模糊+地址保护**：在模糊匹配基础上，增加地址层级检查（省、市、区、街道、路等），避免跨级误配。
    - 未启用的条件请保持“— 不启用 —”。
    """)

    fuzzy_threshold = st.slider("模糊匹配相似度阈值（%）", min_value=40, max_value=100, value=60, step=1,
                                help="只有相似度超过该值才视为匹配。建议先设60%，根据结果微调。")
    enable_address_protect = st.checkbox("启用地址层级保护（推荐，可减少跨区误匹配）", value=True,
                                         help="自动提取省、市、区、街道、路等关键词，若同层级存在不同值则拒绝匹配。")
    show_details = st.checkbox("显示匹配详情（相似度、地址保护是否通过）", value=True,
                               help="在下载结果中增加两列，方便判断阈值是否合适。")

    # ========== 工具函数 ==========
    def clean_str(s):
        """清洗字符串：转大写，移除所有空白字符（包括全角空格、制表符、换行等）"""
        if pd.isna(s):
            return ''
        s = str(s).strip()
        s = re.sub(r'\s+', '', s)
        s = s.replace('\u3000', '')
        return s.upper()

    def remove_brackets(s):
        """删除中英文括号及其中的内容（强制自动执行）"""
        if pd.isna(s):
            return ''
        s = str(s)
        s = re.sub(r'[（(【\[].*?[)）\]】]', '', s)
        return s.strip()

    def split_if_multi(value):
        """拆分多值单元格（用于精确匹配）"""
        if pd.isna(value):
            return ['']
        if re.search(r'[，,;；、\s|/]+', str(value)):
            parts = re.split(r'[，,;；、\s|/]+', str(value))
            parts = [p.strip() for p in parts if p.strip()]
            return parts if parts else [str(value)]
        else:
            return [str(value)]

    def extract_level_keys(text):
        """提取地址中的省、市、区、街道、路等层级的主体关键词"""
        text = clean_str(remove_brackets(text))  # 先清洗并移除括号
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
                if not list1 & list2:
                    return False
        return True

    # ========== 单对地址测试（诊断工具） ==========
    with st.expander("🔧 单对地址测试（诊断隐藏字符或相似度）"):
        test_addr1 = st.text_input("地址A（我的表格中的地址）")
        test_addr2 = st.text_input("地址B（客户表格中的地址）")
        if st.button("测试这对地址"):
            if test_addr1 and test_addr2:
                cleaned1 = clean_str(remove_brackets(test_addr1))
                cleaned2 = clean_str(remove_brackets(test_addr2))
                st.write(f"清洗后地址A：`{cleaned1}`")
                st.write(f"清洗后地址B：`{cleaned2}`")
                if cleaned1 == cleaned2:
                    st.success("✅ 完全相等，应能匹配（即使阈值100也能通过）")
                scores = [
                    fuzz.WRatio(cleaned1, cleaned2),
                    fuzz.partial_ratio(cleaned1, cleaned2),
                    fuzz.token_set_ratio(cleaned1, cleaned2)
                ]
                max_score = max(scores)
                st.write(f"相似度：{max_score:.1f}%")
                protect = address_level_check(cleaned1, cleaned2)
                st.write(f"地址层级保护通过：{'是' if protect else '否'}")
                if not protect:
                    st.warning("地址层级保护未通过，可能是某个层级冲突")
            else:
                st.warning("请填写两个地址")

    # ========== 匹配条件设置 ==========
    none_option = "— 不启用 —"
    my_cols = [none_option] + list(my_df.columns)
    cust_cols = [none_option] + list(customer_df.columns)

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
                        # 模糊匹配：强制移除括号
                        val = remove_brackets(val)
                        row_data.append(clean_str(val))
                my_rows.append(row_data)

            matched_indices = []
            match_details = []

            for cust_idx, cust_row in customer_df.iterrows():
                is_match = False
                best_min_score = 0
                best_protect_pass = True
                matched_my_idx = None

                for my_idx, my_row in enumerate(my_rows):
                    all_pass = True
                    min_score = 100
                    protect_pass = True
                    for cond_idx, (my_col, cust_col, mtype) in enumerate(active_conditions):
                        cust_val = clean_str(cust_row[cust_col])
                        if mtype == "精确":
                            if cust_val not in my_row[cond_idx]:
                                all_pass = False
                                break
                        else:
                            my_val = my_row[cond_idx]

                            # 地址保护
                            if enable_address_protect and mtype in ["模糊", "模糊+地址保护"]:
                                if not address_level_check(my_val, cust_val):
                                    all_pass = False
                                    protect_pass = False
                                    break

                            # 强制移除客户表格中的括号
                            cust_val = remove_brackets(cust_val)
                            my_val = clean_str(my_val)
                            cust_val = clean_str(cust_val)

                            # **完全相等直接通过**
                            if my_val and cust_val and my_val == cust_val:
                                score = 100
                            else:
                                scores = [
                                    fuzz.WRatio(my_val, cust_val),
                                    fuzz.partial_ratio(my_val, cust_val),
                                    fuzz.token_set_ratio(my_val, cust_val)
                                ]
                                score = max(scores)

                            if score < min_score:
                                min_score = score

                            effective_threshold = fuzzy_threshold + (5 if mtype == "模糊+地址保护" else 0)
                            if score < effective_threshold:
                                all_pass = False
                                break

                    if all_pass:
                        if min_score > best_min_score:
                            best_min_score = min_score
                            best_protect_pass = protect_pass
                            matched_my_idx = my_idx
                        is_match = True
                        break

                if is_match:
                    matched_indices.append(cust_idx)
                    match_details.append({
                        '客户行号': cust_idx,
                        '我的表格匹配行号': matched_my_idx,
                        '最小相似度': round(best_min_score, 1),
                        '地址保护通过': '是' if best_protect_pass else '否'
                    })

            matched_df = customer_df.loc[matched_indices].copy()

            st.success(f"✅ 匹配完成！客户表格中共有 {len(matched_df)} 条记录属于你的设备。")

            if len(matched_df) > 0:
                if show_details:
                    details_df = pd.DataFrame(match_details)
                    matched_df = pd.concat([matched_df.reset_index(drop=True), details_df.reset_index(drop=True)], axis=1)

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
                st.warning("没有找到匹配的设备。请尝试降低相似度阈值，或取消地址层级保护，并检查匹配条件设置。")
