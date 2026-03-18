import streamlit as st
import pandas as pd
import docx
from docx.shared import RGBColor
import openpyxl
from openpyxl.styles import Font
from pptx import Presentation
from pptx.dml.color import RGBColor as PptxRGBColor
import os
import re
from io import BytesIO
import pdfplumber

# ページ設定
st.set_page_config(page_title="川内JC 統一ルール修正ツール", page_icon="⚓", layout="wide")
st.title("⚓ 2026年度 川内JC 統一ルール 修正システム")

# 1. 色の設定
color_option = st.selectbox("修正箇所の文字色（Word/Excel/PPT用）", ["赤", "青", "緑", "オレンジ"], index=0)
color_map = {"赤": (255, 0, 0), "青": (0, 0, 255), "緑": (0, 128, 0), "オレンジ": (255, 165, 0)}
selected_rgb = color_map[color_option]

# 2. ルールの読み込み
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        # UTF-8 または Shift-JIS で読み込み
        try:
            df = pd.read_csv('rules.csv', encoding='utf-8')
        except:
            df = pd.read_csv('rules.csv', encoding='shift-jis')
        
        df['len'] = df['類義語'].astype(str).str.len()
        df = df.sort_values('len', ascending=False)
        return dict(zip(df['類義語'].astype(str), df['統一語句'].astype(str)))
    return {}

rules_dict = load_rules()

# 3. 修正ロジック
def apply_rules_to_text(target_text, rules):
    segments = [(target_text, False)]
    for wrong, right in rules.items():
        if wrong == right: continue
        new_segments = []
        for text, already_fixed in segments:
            if already_fixed or wrong not in text:
                new_segments.append((text, already_fixed))
                continue
            
            # 分割して置換
            parts = text.split(wrong)
            for i, part in enumerate(parts):
                if part != "":
                    new_segments.append((part, False))
                if i < len(parts) - 1:
                    new_segments.append((right, True))
        segments = new_segments
    return segments

# --- 各種ファイル修正用関数 ---
def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        parts = apply_rules_to_text(para.text, rules)
        if any(p[1] for p in parts):
            para.text = ""
            for text, is_fixed in parts:
                run = para.add_run(text)
                if is_fixed:
                    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
                    run.bold = True
    out_io = BytesIO()
    doc.save(out_io)
    return out_io.getvalue()

def repair_xlsx(file, rules, rgb):
    wb = openpyxl.load_workbook(file)
    hex_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    parts = apply_rules_to_text(cell.value, rules)
                    if any(p[1] for p in parts):
                        cell.value = "".join([p[0] for p in parts])
                        cell.font = Font(color=hex_color, bold=True)
    out_io = BytesIO()
    wb.save(out_io)
    return out_io.getvalue()

def repair_pptx(file, rules, rgb):
    prs = Presentation(file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    combined_text = "".join(run.text for run in paragraph.runs)
                    parts = apply_rules_to_text(combined_text, rules)
                    if any(p[1] for p in parts):
                        paragraph.text = ""
                        for text, is_fixed in parts:
                            new_run = paragraph.add_run()
                            new_run.text = text
                            if is_fixed:
                                new_run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# --- PDFチェック用関数 ---
def check_pdf(file, rules):
    results = []
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                for wrong, right in rules.items():
                    if wrong != right and wrong in text:
                        # 見つかった箇所を抽出
                        matches = re.finditer(re.escape(wrong), text)
                        for m in matches:
                            start = max(0, m.start() - 10)
                            end = min(len(text), m.end() + 10)
                            context = text[start:end].replace("\n", " ")
                            results.append({
                                "ページ": i + 1,
                                "NGワード": wrong,
                                "修正案": right,
                                "周辺の文章": f"...{context}..."
                            })
    return results

# 4. メイン処理
uploaded_files = st.file_uploader("ファイルをアップロード (Word, Excel, PPT, PDF)", type=["docx", "xlsx", "pptx", "pdf"], accept_multiple_files=True)

if uploaded_files:
    for idx, file in enumerate(uploaded_files):
        ext = file.name.split('.')[-1].lower()
        
        if ext == "pdf":
            with st.spinner(f"PDF {file.name} をスキャン中..."):
                pdf_results = check_pdf(file, rules_dict)
                st.subheader(f"📑 {file.name} のチェック結果")
                if pdf_results:
                    st.warning(f"以下の箇所が統一ルールに抵触している可能性があります。")
                    st.table(pd.DataFrame(pdf_results))
                else:
                    st.success("統一ルールに反する箇所は見つかりませんでした！")
        
        else:
            with st.spinner(f"{file.name} を処理中..."):
                if ext == "docx": data = repair_docx(file, rules_dict, selected_rgb)
                elif ext == "xlsx": data = repair_xlsx(file, rules_dict, selected_rgb)
                elif ext == "pptx": data = repair_pptx(file, rules_dict, selected_rgb)
                
                st.download_button(
                    label=f"📥 修正済みの {file.name} を保存", 
                    data=data, 
                    file_name=f"【修正済】{file.name}",
                    key=f"dl_btn_{file.name}_{idx}"
                )
