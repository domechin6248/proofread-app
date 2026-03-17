import streamlit as st
import pandas as pd
import docx
import openpyxl
from pptx import Presentation
import pdfplumber
import os
from io import BytesIO

# 1. ページ設定
st.set_page_config(
    page_title="2026年度 川内JC 統一ルール校正・修正ツール",
    page_icon="⚓",
    layout="wide"
)

# 2. タイトル
st.title("⚓ 2026年度 川内JC 統一ルール 修正システム")
st.write("ファイルをアップロードすると、自動で修正案を適用したファイルをダウンロードできます。")

# 3. ルールの読み込み
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        return dict(zip(df['類義語'], df['統一語句']))
    else:
        st.error("rules.csv が見つかりません。")
        return {}

rules_dict = load_rules()

# 4. 修正処理の定義
def repair_docx(file, rules):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        for wrong, right in rules.items():
            if str(wrong) in para.text:
                # 書式を維持するために、各ラン（文字の塊）ごとに置換
                for run in para.runs:
                    run.text = run.text.replace(str(wrong), str(right))
    out_io = BytesIO()
    doc.save(out_io)
    return out_io.getvalue()

def repair_xlsx(file, rules):
    wb = openpyxl.load_workbook(file)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for wrong, right in rules.items():
                        cell.value = cell.value.replace(str(wrong), str(right))
    out_io = BytesIO()
    wb.save(out_io)
    return out_io.getvalue()

def repair_pptx(file, rules):
    prs = Presentation(file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for wrong, right in rules.items():
                            run.text = run.text.replace(str(wrong), str(right))
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# 5. メイン処理
uploaded_files = st.file_uploader(
    "ファイルをドロップしてください", 
    type=["docx", "xlsx", "pptx"], # PDFは編集不可のため除外
    accept_multiple_files=True
)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        
        with st.status(f"🛠 {file.name} を修正中...", expanded=False):
            if ext == "docx":
                repaired_data = repair_docx(file, rules_dict)
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == "xlsx":
                repaired_data = repair_xlsx(file, rules_dict)
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == "pptx":
                repaired_data = repair_pptx(file, rules_dict)
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            
            st.write("修正が完了しました。")
            st.download_button(
                label=f"📥 修正済みの {file.name} を保存",
                data=repaired_data,
                file_name=f"【修正済】{file.name}",
                mime=mime,
                key=file.name
            )

    st.info("※PDFは技術的に「上書き修正」ができないため、Word等で修正してからPDF化することをお勧めします。")
