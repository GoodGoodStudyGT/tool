import io
import re

import pandas as pd
import streamlit as st

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ========== 标准字段与别名 ==========
# 标准列名 -> 源表里可能出现的表头别名
STD_ALIASES = {
    '姓名': ['姓名', '名字', '全名', '参检人', 'name', 'Name', 'NAME'],
    '证件号': ['证件号', '证件号码', '身份证', '身份证号', '身份证号码', 'IDNumber', 'idnumber'],
    '手机号': ['手机号', '手机', '电话', '电话号码', '联系电话', 'mobile', 'Mobile', 'phone', 'Phone'],
    '所属组织/部门': ['所属组织/部门', '部门', '组织', '组织/部门', '所属部门', '单位', '单位或就业形态', '所属单位'],
    '体检卡号': ['体检卡号', '卡号', '体检号'],
}

# 由证件号推断的衍生列（模板里也可能列出）
DERIVED_COLS = ['证件类型', '性别']

# 别名（小写、去空格）-> 标准列名，用于模板列名解析
_ALIAS_TO_STD = {}
for _std, _aliases in STD_ALIASES.items():
    for _a in _aliases:
        _ALIAS_TO_STD[_a.strip().lower()] = _std
for _d in DERIVED_COLS:
    _ALIAS_TO_STD[_d.lower()] = _d


def resolve_std(col) -> str:
    """把任意列名解析成标准列名；解析不到就返回去空格后的原名。"""
    key = str(col).strip()
    return _ALIAS_TO_STD.get(key.lower(), key)


# ========== 文本/值清洗 ==========
def clean_text(v) -> str:
    """单个单元格值 -> 干净字符串；缺失值（None / NaN / pd.NA）-> ''。"""
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def get_series(df: pd.DataFrame, col) -> pd.Series:
    """取一列；当表头重复时 df[col] 会是 DataFrame，取其第一列。"""
    s = df[col]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    return s


def clean_series(s: pd.Series) -> pd.Series:
    """整列清洗成普通 Python 字符串（缺失 -> ''）。

    刻意用纯 Python 而非 Series.str：pandas 3.x 默认 pyarrow 字符串后端，
    缺失值是 float nan，astype(str) 不会变成 ''，会污染后续判断。
    """
    return pd.Series([clean_text(v) for v in s.tolist()], index=s.index, dtype=object)


# ========== 证件号解析 ==========
def parse_doc_type(id_number) -> str:
    """粗略判断证件类型"""
    s = clean_text(id_number)
    if re.fullmatch(r"\d{15}", s) or re.fullmatch(r"\d{17}[\dXx]", s):
        return '身份证'
    if re.fullmatch(r"[A-Za-z][0-9]{7,8}", s):
        return '护照'
    if re.fullmatch(r"[HMhm][0-9]{7,9}", s):
        return '港澳台通行证'
    return ''


def parse_gender(id_number) -> str:
    """仅对身份证有效：15 位看最后一位、18 位看倒数第二位，奇数男、偶数女。

    非身份证（护照 / 港澳台 / 空白 / 其它编号）一律返回空串，绝不乱猜性别。
    """
    s = clean_text(id_number).upper()
    if re.fullmatch(r"\d{15}", s):          # 一代证：性别位 = 最后一位
        digit = s[-1]
    elif re.fullmatch(r"\d{17}[\dX]", s):   # 二代证：性别位 = 第 17 位（倒数第二位）
        digit = s[-2]
    else:
        return ''
    try:
        return '男' if int(digit) % 2 else '女'
    except ValueError:
        return ''


# ========== 表头 / 列检测 ==========
def detect_header_row(raw: pd.DataFrame, max_rows: int = 8) -> int:
    """检测真正的表头行：统计该行中“恰好等于”某个关键词的单元格个数。

    用整格精确匹配而不是整行拼接子串匹配，避免把含“身份证”三个字的
    标题行（如“员工身份证信息登记表”）误判成表头。
    """
    keywords = {
        '姓名', '名字', '全名', '参检人',
        '证件号', '证件号码', '身份证', '身份证号', '身份证号码',
        '手机号', '手机', '电话', '电话号码', '联系电话',
        '部门', '单位', '组织', '所属部门', '所属组织/部门',
        '体检卡号', '卡号',
    }
    best_i, best_hits = 0, 0
    for i in range(min(max_rows, len(raw))):
        cells = [clean_text(v) for v in raw.iloc[i].tolist()]
        hits = sum(1 for c in cells if c in keywords)
        if hits >= 2:
            return i
        if hits > best_hits:
            best_i, best_hits = i, hits
    return best_i


def detect_column_by_pattern(df: pd.DataFrame, pattern: str,
                             threshold: float = 0.5, exclude_keywords=()):
    """用纯 Python re 在所有列里找匹配比例最高的列。

    用 re 而不是 Series.str.match：pandas 3.x 的 pyarrow 字符串后端用 RE2，
    不支持 ``\\uXXXX`` 这类转义（会抛 ArrowInvalid: invalid escape sequence: \\u）。
    返回匹配比例最高且达到阈值的列；表头命中 exclude_keywords 的列直接跳过。
    """
    regex = re.compile(pattern)
    best_col, best_ratio = None, 0.0
    for col in df.columns:
        header = str(col)
        if any(k in header for k in exclude_keywords):
            continue
        vals = [v for v in (clean_text(x) for x in get_series(df, col).tolist()) if v]
        if not vals:
            continue
        ratio = sum(1 for v in vals if regex.match(v)) / len(vals)
        if ratio >= threshold and ratio > best_ratio:
            best_col, best_ratio = col, ratio
    return best_col


# ========== 数据读取和字段抽取 ==========
def read_source(data: bytes) -> pd.DataFrame:
    """读取原始 Excel（bytes），自动识别表头，并把表头标签统一成字符串。"""
    raw = pd.read_excel(io.BytesIO(data), header=None, dtype=str)
    hdr = detect_header_row(raw)
    df = pd.read_excel(io.BytesIO(data), header=hdr, dtype=str)
    # read_excel(dtype=str) 只转换“单元格值”，不转换“表头标签”，
    # 数字/日期表头会是 int/float/datetime，下面 .strip() 会崩，先统一成字符串。
    df.columns = [str(c).strip() for c in df.columns]
    return df


def extract_fields(df: pd.DataFrame) -> pd.DataFrame:
    """仅保留目标字段，且总是用身份证推断性别"""
    df = df.rename(columns=lambda c: str(c).strip())
    # 丢掉源表自带的“性别”列（一律按证件号重新推断）
    df = df.drop(columns=[c for c in df.columns if str(c).strip() == '性别'], errors='ignore')

    n = len(df)
    data = {}
    for std, aliases in STD_ALIASES.items():
        col = next((a for a in aliases if a in df.columns), None)
        data[std] = (clean_series(get_series(df, col)) if col is not None
                     else pd.Series([''] * n, index=df.index, dtype=object))

    # 必要字段缺失时，用正则在所有列里猜（取匹配比例最高、且排除明显不相关的表头）
    if data['证件号'].eq('').all():
        c = detect_column_by_pattern(
            df, r"^\d{15}$|^\d{17}[\dXx]$",
            exclude_keywords=('社保', '卡号', '工号', '体检'))
        if c is not None:
            data['证件号'] = clean_series(get_series(df, c))

    if data['手机号'].eq('').all():
        c = detect_column_by_pattern(
            df, r"^1[3-9]\d{9}$",
            exclude_keywords=('卡号', '工号', '社保'))
        if c is not None:
            data['手机号'] = clean_series(get_series(df, c))

    if data['姓名'].eq('').all():
        c = detect_column_by_pattern(
            df, r"^[一-龥·]{2,6}$",
            exclude_keywords=('部门', '组织', '单位', '机构', '城市', '地区',
                              '省', '市', '区', '县', '地址', '所属', '民族', '籍贯'))
        if c is not None:
            data['姓名'] = clean_series(get_series(df, c))

    # 组装 & 追加衍生列（此时各列已是干净字符串，parse_* 不会收到 nan）
    out = pd.DataFrame(data)
    out['证件类型'] = out['证件号'].map(parse_doc_type)
    out['性别'] = out['证件号'].map(parse_gender)        # 只对有效身份证给出男/女
    return out


# ========== 工具函数 ==========
def reorder_columns(df: pd.DataFrame, cols_order: list) -> pd.DataFrame:
    """按给定顺序输出列。模板列名会按别名解析到标准列，
    避免模板把“证件号”写成“身份证号”时数据被整列清空。"""
    out = {}
    for col in cols_order:
        std = resolve_std(col)
        if std in df.columns:
            out[col] = df[std].to_numpy()
        elif str(col).strip() in df.columns:
            out[col] = df[str(col).strip()].to_numpy()
        else:
            out[col] = [''] * len(df)
    return pd.DataFrame(out, index=df.index)


def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


DEFAULT_COLS = ['姓名', '证件类型', '证件号', '手机号', '所属组织/部门', '体检卡号']


@st.cache_data(show_spinner="正在处理 Excel …")
def process(src_bytes: bytes, tpl_bytes):
    """一次性完成解析、拆分、生成下载字节；按上传内容缓存，避免每次重跑。"""
    df_all = extract_fields(read_source(src_bytes))
    cols = (pd.read_excel(io.BytesIO(tpl_bytes), nrows=0).columns.tolist()
            if tpl_bytes else DEFAULT_COLS)

    male = reorder_columns(df_all[df_all['性别'] == '男'], cols)
    female = reorder_columns(df_all[df_all['性别'] == '女'], cols)
    unknown = reorder_columns(df_all[~df_all['性别'].isin(['男', '女'])], cols)

    return {
        'counts': (len(df_all), len(male), len(female), len(unknown)),
        'unknown_preview': unknown,
        'male_bytes': df_to_xlsx_bytes(male),
        'female_bytes': df_to_xlsx_bytes(female),
        'unknown_bytes': df_to_xlsx_bytes(unknown),
    }


# ========== Streamlit 前端 ==========
st.set_page_config(page_title="Excel 智能提取 & 性别拆分", page_icon="📑", layout="centered")
st.title("📑 Excel 智能提取 & 性别拆分 工具")
st.markdown(
    """
上传一个含 **“姓名 / 证件号 / 手机号 / 所属组织 / 体检卡号”** 等字段的 Excel
系统自动识别表头，多表头也支持。处理后按身份证号自动拆分男女，并提供 XLSX 下载。
> 性别只根据**有效的 15/18 位身份证**推断；护照、港澳台证件、空白或异常证件号会归入「未识别」，不会被随意分到男女里。
"""
)

src_file = st.file_uploader("⬆️ 上传待处理 Excel", type=["xls", "xlsx"])
tpl_file = st.file_uploader("⬆️ 上传模板文件（可选：决定输出列顺序）", type=["xls", "xlsx"])

if src_file:
    result = process(src_file.getvalue(), tpl_file.getvalue() if tpl_file else None)
    total, n_male, n_female, n_unknown = result['counts']

    st.success(f"✅ 处理完成：共 {total} 行 → 男 {n_male} / 女 {n_female} / 未识别 {n_unknown}")

    if n_unknown:
        st.warning(
            f"⚠️ 有 {n_unknown} 行无法从证件号识别性别"
            "（证件号为空 / 非 15或18 位身份证 / 格式异常），未计入男女文件，请人工核对："
        )
        st.dataframe(result['unknown_preview'], use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇️ 下载男性结果 (XLSX)", result['male_bytes'],
                           file_name="male.xlsx", mime=XLSX_MIME)
    with col2:
        st.download_button("⬇️ 下载女性结果 (XLSX)", result['female_bytes'],
                           file_name="female.xlsx", mime=XLSX_MIME)

    if n_unknown:
        st.download_button("⬇️ 下载未识别 / 待人工核对 (XLSX)", result['unknown_bytes'],
                           file_name="unknown.xlsx", mime=XLSX_MIME)
else:
    st.info("👈 请先上传源 Excel 文件")
