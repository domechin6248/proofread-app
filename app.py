import streamlit as st
import pandas as pd
import docx
import openpyxl
from pptx import Presentation
import pdfplumber
import os

# 1. ページ設定
st.set_page_config(
    page_title="2026年度 川内JC 統一ルール校正ツール",
    page_icon="⚓",
    layout="wide"
)

# 2. タイトルと説明（サイドバーなしでスッキリさせました）
st.title("⚓ 2026年度 川内JC 統一ルール校正システム")
st.write("Word, Excel, PowerPoint, PDFをドロップして一括チェックできます。")

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

# 4. ファイルアップロード
uploaded_files = st.file_uploader(
    "チェックしたいファイルをドロップしてください（複数可）", 
    type=["docx", "xlsx", "pptx", "pdf"], 
    accept_multiple_files=True
)

def check_text(text, filename, location):
    found = []
    if not isinstance(text, str):
        return found
    for wrong, right in rules_dict.items():
        if str(wrong) in text:
            found.append({
                "ファイル名": filename, 
                "箇所": location, 
                "指摘内容": f"「{wrong}」が含まれています",
                "修正案": f"「{right}」に統一してください"
            })
    return found

# 5. 解析処理
if uploaded_files:
    all_errors = []
    progress_bar = st.progress(0)
    
    for idx, file in enumerate(uploaded_files):
        ext = file.name.split('.')[-1].lower()
        
        if ext == "docx":
            doc = docx.Document(file)
            for i, para in enumerate(doc.paragraphs):
                all_errors.extend(check_text(para.text, file.name, f"{i+1}行目"))
        
        elif ext == "xlsx":
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet in wb.worksheets:
                for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                    for cell_idx, cell_value in enumerate(row):
                        if cell_value:
                            all_errors.extend(check_text(str(cell_value), file.name, f"シート:{sheet.title} ({row_idx+1}行目)"))

        elif ext == "pptx":
            prs = Presentation(file)
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        all_errors.extend(check_text(shape.text, file.name, f"{i+1}枚目スライド"))

        elif ext == "pdf":
            with pdfplumber.open(file) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_errors.extend(check_text(text, file.name, f"{i+1}ページ目"))
        
        progress_bar.progress((idx + 1) / len(uploaded_files))

    st.divider()
    if all_errors:
        st.warning(f"⚠️ {len(all_errors)} 件の修正推奨箇所が見つかりました。")
        st.table(pd.DataFrame(all_errors))
    else:
        st.success("✨ チェック完了！すべての資料で統一ルールが守られています。")
