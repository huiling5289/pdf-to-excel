import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# 设置页面标题和图标
st.set_page_config(page_title="PDF 批量轉 Excel 工具", page_icon="📄", layout="centered")

# 主标题
st.markdown('<h1 style="color:#0000FF;">📄 PDF 批量轉 Excel 工具</h1>', unsafe_allow_html=True)  # 蓝色主标题
st.markdown("### 支持上傳多個 PDF 文件，並轉換為 Excel文件")

# 上传多个 PDF 文件
uploaded_files = st.file_uploader("上傳你的 PDF 文件（支持多選）", type=["pdf"], accept_multiple_files=True)

# 清理數據函數
def clean_data(table):
    cleaned_table = []
    for row in table:
        cleaned_row = []
        for cell in row:
            if isinstance(cell, str):
                cell = cell.replace(",", "").replace("(", "-").replace(")", "")
                try:
                    cleaned_row.append(float(cell))
                except ValueError:
                    cleaned_row.append(cell)
            else:
                cleaned_row.append(cell)
        cleaned_table.append(cleaned_row)
    return cleaned_table

# PDF 轉 Excel 函數
def pdf_to_excel(pdf_file):
    output = BytesIO()
    with pdfplumber.open(pdf_file) as pdf:
        all_data = []
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                st.warning(f"第 {page_number} 頁沒有檢測到表格，跳過...")
                continue
            for table in tables:
                cleaned_table = clean_data(table)
                all_data.extend(cleaned_table)

        df = pd.DataFrame(all_data)
        df.to_excel(output, index=False, header=False)
        output.seek(0)
    return output

# 批量處理上傳的 PDF 文件
if uploaded_files:
    st.success(f"已成功上傳 {len(uploaded_files)} 個文件！")
    for uploaded_file in uploaded_files:
        st.markdown(f"### 文件名稱: {uploaded_file.name}")
        try:
            excel_data = pdf_to_excel(uploaded_file)
            st.download_button(
                label=f"下載 {uploaded_file.name} 的 Excel 文件 📥",
                data=excel_data,
                file_name=f"{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"處理 {uploaded_file.name} 時出現錯誤：{e}")