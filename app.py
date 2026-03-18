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
        try:
            df = pd.read_csv('rules.csv', encoding='utf-8')
        except:
            df = pd.read_csv('rules.csv', encoding='shift-jis')
        
        df['len'] = df['類義語'].astype(str).str.len()
        df = df.sort_values('len', ascending=False)
        return dict(zip(df['類義語'].astype(str), df['統一語句'].astype(str)))
    return {}

rules_dict = load_rules()

# 3. 修正・熟語保護ロジック
def apply_rules_to_text(target_text, rules):
    # 【絶対保護リスト】ここに書かれた定型文や熟語は、ルール適用から除外されます
    keep_words = [
        "会員拡大運動", "正会員", "新入会員",
        # 60周年への誓い
        "会員に成長する機会を提供し、誇りを持てる持続可能な組織運営を行います。",
        "国内だけでなく国際的な視野を持ち、社会貢献を行える組織を目指します。",
        # JC宣言・綱領関連
        "日本の青年会議所は", "希望をもたらす変革の起点として", "輝く個性が調和する未来を描き",
        "社会の課題を解決することで", "持続可能な地域を創ることを誓う",
        "われわれ JAYCEE は", "社会的・国家的・国際的な責任を自覚し",
        # 志高きビジョン関連
        "市民がまちづくりへの当事者意識を持ち", "多様な市民が共生できるまちを創造します",
        "地域特有の個性を活かし"
    ]
    
    # 1. 保護したい言葉を一時的に避難させる
    protected_text = target_text
    placeholders = {}
    for i, word in enumerate(keep_words):
        if word in protected_text:
            placeholder = f"__KEEP_WORD_{i}__"
            placeholders[placeholder] = word
            protected_text = protected_text.replace(word, placeholder)

    segments = [(protected_text, False)]
    
    # 2. ルールの適用
    for wrong, right in rules.items():
        if wrong == right: continue
        new_segments = []
        for text, already_fixed in segments:
            if already_fixed or wrong not in text:
                new_segments.append((text, already_fixed))
                continue
            
            parts = text.split(wrong)
            for j, part in enumerate(parts):
                if part != "":
                    new_segments.append((part, False))
                if j < len(parts) - 1:
                    new_segments.append((right, True))
        segments = new_segments

    # 3. 避難させていた言葉を元に戻す
    final_segments = []
    for text, is_fixed in segments:
        if not is_fixed:
            temp_text = text
            for placeholder, original in placeholders.items():
                temp_text = temp_text.replace(placeholder, original)
            final_segments.append((temp_text, False))
        else:
            final_segments.append((text, is_fixed))
            
    return final_segments

# --- 各種ファイル修正用関数（中略） ---
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

def check_pdf(file, rules):
    results = []
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                for wrong, right in rules.items():
                    if wrong != right and wrong in text:
                        matches = re.finditer(re.escape(wrong), text)
                        for m in matches:
                            # 前後の文脈を確認し、保護対象が含まれていればスキップ
                            start = max(0, m.start() - 50)
                            end = min(len(text), m.end() + 50)
                            context_full = text[start:end]
                            
                            # apply_rules_to_textと同じ保護ロジックを簡易適用
                            is_protected = False
                            keep_list = ["会員拡大運動", "正会員", "新入会員", "会員に成長する機会を提供し"]
                            for kw in keep_list:
                                if kw in context_full:
                                    is_protected = True
                                    break
                            
                            if not is_protected:
                                results.append({
                                    "ページ": i + 1,
                                    "NGワード": wrong,
                                    "修正案": right,
                                    "周辺の文章": f"...{text[max(0, m.start()-10):min(len(text), m.end()+10)].replace('\n', ' ')}..."
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
                st.download_button(label=f"📥 修正済みの {file.name} を保存", data=data, file_name=f"【修正済】{file.name}", key=f"dl_btn_{file.name}_{idx}")
