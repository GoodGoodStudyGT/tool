import streamlit as st
import tempfile
import shutil
import os
from excell.extract_and_split_template import main

st.title("Excel 智能提取 & 性别拆分")

uploaded = st.file_uploader("📁 上传你的 Excel 文件", type=["xls","xlsx"])
if uploaded:
    # 写入临时文件
    tmp = tempfile.mkdtemp()
    src_path = os.path.join(tmp, uploaded.name)
    with open(src_path, "wb") as f:
        f.write(uploaded.getbuffer())

    # 运行处理函数
    out_dir = os.path.join(tmp, "out")
    main(src_path, template_file=None, output_dir=out_dir)

    # 打包成 ZIP
    zip_path = os.path.join(tmp, "result.zip")
    shutil.make_archive(zip_path.replace(".zip",""), 'zip', out_dir)

    # 提供下载
    with open(zip_path, "rb") as fp:
        st.download_button("⬇️ 下载处理结果", fp, file_name="result.zip")
