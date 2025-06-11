import streamlit as st
import tempfile
import shutil
import os

import sys, os
# 假设 streamlit_app.py 和 excell/ 在同一目录
sys.path.insert(0, os.path.dirname(__file__))


from extract_and_split_template import main

st.title("Excel 智能提取 & 性别拆分")

uploaded = st.file_uploader("📁 上传你的 Excel 文件", type=["xls","xlsx"])
if uploaded:
    # 写入临时源文件
    tmpdir = tempfile.mkdtemp()
    src = os.path.join(tmpdir, uploaded.name)
    with open(src, "wb") as f: f.write(uploaded.getbuffer())

    # 模板文件就直接用仓库里的 template.xlsx
    repo_root = os.getcwd()
    tpl = os.path.join(repo_root, "template.xlsx")  # 确保这个文件已提交到 GitHub

    # 处理
    out = os.path.join(tmpdir, "out")
    main(src, template_file=tpl, output_dir=out)
    
    # 打包并下载
    zipf = os.path.join(tmpdir, "result")
    shutil.make_archive(zipf, 'zip', out)
    with open(zipf + ".zip", "rb") as fp:
        st.download_button("⬇️ 下载结果", fp, file_name="result.zip")
        
