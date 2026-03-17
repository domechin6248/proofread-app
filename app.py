import streamlit as st
import pandas as pd
import docx
import openpyxl
from pptx import Presentation
import pdfplumber
import os

# 1. ページ設定（ブラウザのタブに表示される情報）
st.set_page_config(
    page_title="川内JC 2026年度 統一ルール校正ツール",
    page_icon="⚓",
    layout="wide"
)

# 2. サイドバーの設定（監事メッセージ）
with st.sidebar:
    st.header("⚓ 監事の視点")
    st.info("「正確な語句の使用は、組織の信頼に直結します。一文字の妥協が、組織の品格を左右することを忘れないでください。」")
    st.write("---")
    st.caption("2026年度 監事監修ツール")
    st.write("対象：Word, Excel, PPT, PDF")

# 3. タイトルと説明
st.title("⚓ 2026年度 川内JC 統一ルール校正システム")
st.markdown("### ～ 薩摩川内の未来を創る、正確な資料作りを ～")
st.write("資料をアップロードすると、統一語句や書式を自動で一括チェックします。")

# 4. ルールの読み込み（キャッシュ機能で高速化）
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        # 1行目の項目名が「類義語」「統一語句」であることを想定
        return dict(zip(df['類義語'], df['統一語句']))
    else:
        st.error("rules.csv が見つかりません。GitHubのトップページに配置してください。")
        return {}

rules_dict = load_rules()

# 5. ファイルアップロード（複数・多形式対応）
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

# 6. 解析処理
if uploaded_files:
    all_errors = []
    progress_bar = st.progress(0)
    
    for idx, file in enumerate(uploaded_files):
        ext = file.name.split('.')[-1].lower()
        
        # Word
        if ext == "docx":
            doc = docx.Document(file)
            for i, para in enumerate(doc.paragraphs):
                all_errors.extend(check_text(para.text, file.name, f"{i+1}行目"))
        
        # Excel
        elif ext == "xlsx":
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet in wb.worksheets:
                for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                    for cell_idx, cell_value in enumerate(row):
                        if cell_value:
                            all_errors.extend(check_text(str(cell_value), file.name, f"シート:{sheet.title} ({row_idx+1}行目)"))

        # PowerPoint
        elif ext == "pptx":
            prs = Presentation(file)
            for i, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        all_errors.extend(check_text(shape.text, file.name, f"{i+1}枚目スライド"))

        # PDF
        elif ext == "pdf":
            with pdfplumber.open(file) as pdf:
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        all_errors.extend(check_text(text, file.name, f"{i+1}ページ目"))
        
        progress_bar.progress((idx + 1) / len(uploaded_files))

    # 7. 結果表示
    st.divider()
    if all_errors:
        st.warning(f"⚠️ {len(all_errors)} 件の修正推奨箇所が見つかりました。")
        st.table(pd.DataFrame(all_errors))
    else:
        st.success("✨ チェック完了！すべての資料で統一ルールが守られています。")
