import os
import re
import tempfile

import pandas as pd
import streamlit as st


# ========== 核心解析函数 ==========
def parse_gender(id_number: str) -> str:
    """身份证倒数第二位：奇数男，偶数女；否则返回空串"""
    if not isinstance(id_number, str) or len(id_number) < 2:
        return ''
    try:
        return '男' if int(id_number[-2]) % 2 else '女'
    except ValueError:
        return ''


def parse_doc_type(id_number: str) -> str:
    """粗略判断证件类型"""
    s = (id_number or '').strip()
    if re.fullmatch(r"\d{15}", s) or re.fullmatch(r"\d{17}[\dXx]", s):
        return '身份证'
    if re.fullmatch(r"[A-Za-z][0-9]{7,8}", s):
        return '护照'
    if re.fullmatch(r"[HMhm][0-9]{7,9}", s) or (s.isdigit() and len(s) in (8, 9, 10)):
        return '港澳台通行证'
    return ''


def detect_header_row(raw: pd.DataFrame, max_rows: int = 5) -> int:
    """检测真正的表头行索引"""
    keywords = ['姓名', '证件号', '身份证', '手机号', '电话']
    for i in range(min(max_rows, len(raw))):
        if any(k in ''.join(raw.iloc[i].fillna('')) for k in keywords):
            return i
    return 0


def detect_column_by_pattern(df: pd.DataFrame, pattern: str, threshold: float = 0.5):
    """用正则猜测符合特征的列"""
    regex = re.compile(pattern)
    for col in df.columns:
        s = df[col].dropna().astype(str)
        if len(s) and (s.str.match(regex).sum() / len(s) >= threshold):
            return col
    return None


# ========== 数据读取和字段抽取 ==========
def read_source(path: str) -> pd.DataFrame:
    """读取原始 Excel，自动识别表头"""
    raw = pd.read_excel(path, header=None, dtype=str)
    hdr = detect_header_row(raw)
    return pd.read_excel(path, header=hdr, dtype=str)


def extract_fields(df: pd.DataFrame) -> pd.DataFrame:
    """仅保留目标字段，且总是用身份证推断性别"""
    # 1) 丢掉源表自带的“性别”列
    df = df.drop(columns=[c for c in df.columns if c.strip() == '性别'], errors='ignore')

    # 2) 明确映射
    mapping = {
        '姓名': ['姓名', 'name', 'Name', '参检人'],
        '证件号': ['证件号', '身份证号', '身份证号码', 'IDNumber'],
        '手机号': ['手机号', '电话', '电话号码', 'mobile', 'Mobile'],
        '所属组织/部门': ['所属组织/部门', '部门', '组织', '组织/部门', '单位', '单位或就业形态'],
        '体检卡号': ['体检卡号', '卡号'],
    }

    data = {}
    for std, aliases in mapping.items():
        col = next((a for a in aliases if a in df.columns), None)
        data[std] = df[col].astype(str) if col else pd.Series([''] * len(df), index=df.index)

    # 3) 若必要字段缺失，再用正则去猜
    if data['证件号'].eq('').all():
        c = detect_column_by_pattern(df, r"^\d{15}$|^\d{17}[\dXx]$")
        if c:
            data['证件号'] = df[c].astype(str)

    if data['手机号'].eq('').all():
        c = detect_column_by_pattern(df, r"^1[3-9]\d{9}$")
        if c:
            data['手机号'] = df[c].astype(str)

    if data['姓名'].eq('').all():
        c = detect_column_by_pattern(df, r"^[\u4e00-\u9fa5]{2,4}$")
        if c:
            data['姓名'] = df[c].astype(str)

    # 4) 组装 & 追加衍生列
    out = pd.DataFrame(data, dtype=str)
    out['证件类型'] = out['证件号'].apply(parse_doc_type)
    out['性别'] = out['证件号'].apply(parse_gender)        # 一律按身份证算

    return out


# ========== 工具函数 ==========
def reorder_columns(df: pd.DataFrame, cols_order: list) -> pd.DataFrame:
    """按给定顺序保留/补全列，其他多余列全部忽略"""
    return df.reindex(columns=cols_order, fill_value='')


DEFAULT_COLS = ['姓名', '证件类型', '证件号', '手机号', '所属组织/部门', '体检卡号']


# ========== Streamlit 前端 ==========
st.set_page_config(page_title="Excel 智能提取 & 性别拆分", page_icon="📑", layout="centered")
st.title("📑 Excel 智能提取 & 性别拆分 工具")
st.markdown(
    """
上传一个含 **“姓名 / 证件号 / 手机号 / 所属组织 / 体检卡号”** 等字段的 Excel  
系统自动识别表头，多表头也支持。处理后按身份证号自动拆分男女，并提供两个 XLSX 下载。
"""
)

src_file = st.file_uploader("⬆️ 上传待处理 Excel", type=["xls", "xlsx"])
tpl_file = st.file_uploader("⬆️ 上传模板文件（可选：决定输出列顺序）", type=["xls", "xlsx"])

if src_file:
    tmpdir = tempfile.mkdtemp()
    src_path = os.path.join(tmpdir, src_file.name)
    with open(src_path, "wb") as f:
        f.write(src_file.getbuffer())

    tpl_path = None
    if tpl_file:
        tpl_path = os.path.join(tmpdir, tpl_file.name)
        with open(tpl_path, "wb") as f:
            f.write(tpl_file.getbuffer())

    # ---------- 主逻辑 ----------
    df_raw = read_source(src_path)
    df_all = extract_fields(df_raw)

    # 输出列顺序：模板优先
    if tpl_path:
        tpl_cols = pd.read_excel(tpl_path, nrows=0).columns.tolist()
        male_df = reorder_columns(df_all[df_all['性别'] == '男'], tpl_cols)
        female_df = reorder_columns(df_all[df_all['性别'] == '女'], tpl_cols)
    else:
        male_df = reorder_columns(df_all[df_all['性别'] == '男'], DEFAULT_COLS)
        female_df = reorder_columns(df_all[df_all['性别'] == '女'], DEFAULT_COLS)

    # ---------- 写文件并供下载 ----------
    outdir = os.path.join(tmpdir, "out")
    os.makedirs(outdir, exist_ok=True)
    male_path = os.path.join(outdir, "male.xlsx")
    female_path = os.path.join(outdir, "female.xlsx")

    male_df.to_excel(male_path, index=False)
    female_df.to_excel(female_path, index=False)

    st.success("✅ 处理完成！请点击下方按钮下载结果")

    with open(male_path, "rb") as mfile:
        st.download_button(
            "⬇️ 下载男性结果 (XLSX)",
            mfile,
            file_name="male.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with open(female_path, "rb") as ffile:
        st.download_button(
            "⬇️ 下载女性结果 (XLSX)",
            ffile,
            file_name="female.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("👈 请先上传源 Excel 文件")
