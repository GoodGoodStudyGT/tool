import pandas as pd
import os
import re
import tempfile
import shutil
import streamlit as st

# --------- 核心逻辑函数 ---------
def parse_gender(id_number: str) -> str:
    """
    根据身份证倒数第二位判断性别，奇数男、偶数女。
    """
    if not isinstance(id_number, str) or len(id_number) < 2:
        return ''
    try:
        digit = int(id_number[-2])
    except Exception:
        return ''
    return '男' if digit % 2 else '女'


def parse_doc_type(id_number: str) -> str:
    """
    简单推断证件类型：身份证、护照、港澳台通行证。
    """
    s = (id_number or '').strip()
    if re.match(r"^\d{15}$", s) or re.match(r"^\d{17}[\dXx]$", s):
        return '身份证'
    if re.match(r"^[A-Za-z][0-9]{7,8}$", s):
        return '护照'
    if re.match(r"^[HMhm][0-9]{7,9}$", s) or (s.isdigit() and len(s) in (8,9,10)):
        return '港澳台通行证'
    return ''


def detect_header_row(raw: pd.DataFrame, max_rows: int = 5) -> int:
    keywords = ['姓名','证件号','身份证','手机号','电话']
    for i in range(min(max_rows, len(raw))):
        text = ''.join(raw.iloc[i].fillna(''))
        if any(kw in text for kw in keywords):
            return i
    return 0


def detect_column_by_pattern(df: pd.DataFrame, pattern: str, threshold: float = 0.5) -> str:
    regex = re.compile(pattern)
    for col in df.columns:
        s = df[col].dropna().astype(str)
        if len(s) == 0:
            continue
        if s.str.match(regex).sum() / len(s) >= threshold:
            return col
    return None


def read_source(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    hdr = detect_header_row(raw)
    return pd.read_excel(path, header=hdr, dtype=str)


def extract_fields(df: pd.DataFrame) -> pd.DataFrame:
    # 基于列名的初始映射，包含新增字段
    mapping = {
        '姓名': ['姓名','name','Name','参检人'],
        '证件号': ['证件号','身份证号','IDNumber'],
        '手机号': ['手机号','电话','mobile','Mobile'],
        '所属组织/部门': ['所属组织/部门','部门','组织','组织/部门'],
        '体检卡号': ['体检卡号','卡号','体检卡号']
    }
    data = {}
    for k, cands in mapping.items():
        col = next((c for c in cands if c in df.columns), None)
        data[k] = df[col].astype(str) if col else pd.Series([''] * len(df), index=df.index)

    # 模式补全：姓名、证件号、手机号
    if data['证件号'].eq('').all():
        c = detect_column_by_pattern(df, r"^\d{17}[\dXx]$")
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

    # 构造 DataFrame 并添加证件类型、性别
    df_out = pd.DataFrame(data)
    df_out['证件类型'] = df_out['证件号'].apply(parse_doc_type)
    df_out['性别'] = df_out['证件号'].apply(parse_gender)
    return df_out

# --------- Streamlit 前端 ---------
st.title("📑 Excel 智能提取 & 性别拆分 工具")
st.markdown("上传一个含“姓名”、“证件号”、“手机号”等字段的 Excel，自动识别表头，多表头也支持。处理后按性别拆分并打包下载。结果表列顺序固定，表头不变。")

src_file = st.file_uploader("上传待处理Excel", type=["xls","xlsx"] )
tpl_file = st.file_uploader("上传模板文件（可选）", type=["xls","xlsx"] )

def reorder_columns(df: pd.DataFrame, cols_order: list) -> pd.DataFrame:
    # 只保留需展示的列并按指定顺序排列
    return df.reindex(columns=cols_order, fill_value='')

# 指定默认列顺序
DEFAULT_COLS = ['姓名', '证件类型', '证件号', '手机号', '所属组织/部门', '体检卡号']

if src_file:
    tmpdir = tempfile.mkdtemp()
    # 保存上传的源文件
    src_path = os.path.join(tmpdir, src_file.name)
    with open(src_path, "wb") as f:
        f.write(src_file.getbuffer())
    tpl_path = None
    if tpl_file:
        tpl_path = os.path.join(tmpdir, tpl_file.name)
        with open(tpl_path, "wb") as f:
            f.write(tpl_file.getbuffer())

    # 提取并拆分
    df = read_source(src_path)
    df_all = extract_fields(df)
    # 重排序列
    if tpl_path:
        # 按模板表头顺序
        tpl_cols = pd.read_excel(tpl_path, nrows=0).columns.tolist()
        male_df = reorder_columns(df_all[df_all['性别']=='男'].drop(columns=['性别']), tpl_cols)
        female_df = reorder_columns(df_all[df_all['性别']=='女'].drop(columns=['性别']), tpl_cols)
    else:
        male_df = reorder_columns(df_all[df_all['性别']=='男'].drop(columns=['性别']), DEFAULT_COLS)
        female_df = reorder_columns(df_all[df_all['性别']=='女'].drop(columns=['性别']), DEFAULT_COLS)

    # 写入文件
    outdir = os.path.join(tmpdir, 'out')
    os.makedirs(outdir, exist_ok=True)
    m_path = os.path.join(outdir, 'male.xlsx')
    f_path = os.path.join(outdir, 'female.xlsx')
    male_df.to_excel(m_path, index=False)
    female_df.to_excel(f_path, index=False)

    # 打包下载
    zipname = os.path.join(tmpdir, 'result')
    shutil.make_archive(zipname, 'zip', outdir)
    with open(zipname + '.zip', 'rb') as fp:
        st.download_button("⬇️ 下载结果 (ZIP)", fp, file_name="result.zip")
