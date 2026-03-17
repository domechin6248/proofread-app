import streamlit as st
import pandas as pd
import docx
import openpyxl
from pptx import Presentation
import pdfplumber
import os

st.set_page_config(page_title="2026年度 総合校正システム", layout="wide")
st.title("📋 2026年度 総合資料校正ツール")
st.write("Word, Excel, PPT, PDF を一括チェックします。")

# 1. ルールの読み込み
@st.cache_data
def load_rules():
    df = pd.read_csv('rules.csv')
    return dict(zip(df['類義語'], df['統一語句']))

rules_dict = load_rules()

# 2. ファイルアップロード（複数の形式を許可）
uploaded_files = st.file_uploader("資料をアップロードしてください", 
                                  type=["docx", "xlsx", "pptx", "pdf"], 
                                  accept_multiple_files=True)

def check_text(text, filename, location):
    found = []
    for wrong, right in rules_dict.items():
        if str(wrong) in text:
            found.append({"ファイル": filename, "箇所": location, "内容": f"「{wrong}」→「{right}」"})
    return found

if uploaded_files:
    all_errors = []
    for file in uploaded_files:
        ext = file.name.split('.')[-1]
        
        # Wordのチェック
        if ext == "docx":
            doc = docx.Document(file)
            for i, para in enumerate(doc.paragraphs):
                all_errors.extend(check_text(para.text, file.name, f"{i+1}行目"))
        
        # Excelのチェック
        elif ext == "xlsx":
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell_idx, cell_value in enumerate(row):
                        if cell_value and isinstance(cell_value, str):
                            all_errors.extend(check_text(cell_value, file.name, f"シート:{sheet.title}"))

        # PowerPointのチェック
        elif ext == "pptx":
            prs = Presentation(file)
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        all_errors.extend(check_text(shape.text, file.name, f"{i+1}枚目スライド"))

        # PDFのチェック
        elif ext == "pdf":
            with pdfplumber.open(file) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_errors.extend(check_text(text, file.name, f"{i+1}ページ目"))

    if all_errors:
        st.table(pd.DataFrame(all_errors))
    else:
        st.success("すべての資料に問題ありませんでした！")
