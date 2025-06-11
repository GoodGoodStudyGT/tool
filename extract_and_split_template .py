import pandas as pd
import argparse
import os
import re


def parse_gender(id_number: str) -> str:
    """
    根据身份证号码的倒数第二位数字判断性别，奇数为男，偶数为女。
    """
    if not isinstance(id_number, str) or len(id_number) < 2:
        return ''
    try:
        sex_digit = int(id_number[-2])
    except Exception:
        return ''
    return '男' if sex_digit % 2 else '女'


def parse_doc_type(id_number: str) -> str:
    """
    根据证件号简单匹配推断证件类型：
      - 15 或 18 位数字（尾位可 X/x）: 身份证
      - 护照号: 1 位字母 + 7~8 位数字
      - 港澳台通行证: H/M 开头 + 数字或纯数字 8~10 位
    """
    if not isinstance(id_number, str):
        return ''
    s = id_number.strip()
    # 身份证
    if re.match(r"^\d{15}$", s) or re.match(r"^\d{17}[\dXx]$", s):
        return '身份证'
    # 护照
    if re.match(r"^[a-zA-Z][0-9]{7,8}$", s):
        return '护照'
    # 港澳台通行证
    if re.match(r"^[HMhm][0-9]{7,9}$", s) or (s.isdigit() and len(s) in [8,9,10]):
        return '港澳台通行证'
    return ''


def detect_header_row(raw_df: pd.DataFrame, max_rows: int = 5) -> int:
    keywords = ['姓名', '证件号', '身份证', '手机号', '电话']
    for i in range(min(max_rows, len(raw_df))):
        row_str = ''.join(raw_df.iloc[i].fillna(''))
        if any(kw in row_str for kw in keywords):
            return i
    return 0


def detect_column_by_pattern(df: pd.DataFrame, pattern: str, threshold: float = 0.5) -> str:
    regex = re.compile(pattern)
    for col in df.columns:
        series = df[col].dropna().astype(str)
        if len(series) == 0:
            continue
        match_count = series.str.match(regex).sum()
        if match_count / len(series) >= threshold:
            return col
    return None


def extract_fields(df: pd.DataFrame) -> pd.DataFrame:
    # 基于列名的初始映射
    name_map = {
        '姓名': ['姓名', 'name', 'Name', '参检人', '姓名(全称)'],
        '证件号': ['证件号', '身份证号', 'IDNumber'],
        '手机号': ['手机号', '电话', 'mobile', 'Mobile']
    }
    extracted = {}
    for field, candidates in name_map.items():
        found = next((col for col in candidates if col in df.columns), None)
        extracted[field] = df[found].astype(str) if found else pd.Series([''] * len(df), index=df.index)

    # 数据模式检测补全
    if extracted['证件号'].eq('').all():
        col = detect_column_by_pattern(df, r"^\d{17}[\dXx]$")
        if col:
            extracted['证件号'] = df[col].astype(str)
    if extracted['手机号'].eq('').all():
        col = detect_column_by_pattern(df, r"^1[3-9]\d{9}$")
        if col:
            extracted['手机号'] = df[col].astype(str)
    if extracted['姓名'].eq('').all():
        col = detect_column_by_pattern(df, r"^[\u4e00-\u9fa5]{2,4}$")
        if col:
            extracted['姓名'] = df[col].astype(str)

    # 根据证件号推断证件类型
    extracted['证件类型'] = extracted['证件号'].apply(parse_doc_type)

    result = pd.DataFrame(extracted)
    result['性别'] = result['证件号'].apply(parse_gender)
    return result


def read_source_excel(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    header_row = detect_header_row(raw)
    return pd.read_excel(path, header=header_row, dtype=str)


def read_template_columns(path: str) -> list:
    ext = os.path.splitext(path)[1].lower()
    engine = None
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        engine = 'openpyxl'
    elif ext == '.xls':
        engine = 'xlrd'
    try:
        if engine:
            df = pd.read_excel(path, nrows=0, dtype=str, engine=engine)
        else:
            df = pd.read_excel(path, nrows=0, dtype=str)
    except Exception:
        df = pd.read_excel(path, nrows=0, dtype=str)
    return list(df.columns)


def main(source_file: str, template_file: str = None, output_dir: str = '.') -> None:
    df_src = read_source_excel(source_file)
    df_ext = extract_fields(df_src)
    df_male = df_ext[df_ext['性别'] == '男'].drop(columns=['性别'])
    df_female = df_ext[df_ext['性别'] == '女'].drop(columns=['性别'])

    # 按模板列顺序输出
    if template_file:
        cols = read_template_columns(template_file)
        df_male = df_male.reindex(columns=cols, fill_value='')
        df_female = df_female.reindex(columns=cols, fill_value='')

    os.makedirs(output_dir, exist_ok=True)
    male_path = os.path.join(output_dir, 'template_male.xlsx')
    female_path = os.path.join(output_dir, 'template_female.xlsx')
    df_male.to_excel(male_path, index=False)
    df_female.to_excel(female_path, index=False)
    print(f"已生成：{male_path} 和 {female_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='智能提取并按性别拆分模板字段')
    parser.add_argument('source', help='源 Excel 路径')
    parser.add_argument('-t', '--template', help='模板文件（保持列顺序）', default='temple.xlsx')
    parser.add_argument('-o', '--output', help='输出目录，默认为当前', default='.')
    args = parser.parse_args()
    main(args.source, args.template, args.output)
