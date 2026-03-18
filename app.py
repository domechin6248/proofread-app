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
color_option = st.selectbox("修正箇所の文字色（Word/Excel/PPT用）", ["赤", "青", "緑", "黒"], index=0)
color_map = {"赤": (255, 0, 0), "青": (0, 0, 255), "緑": (0, 128, 0), "黒": (0, 0, 0)}
selected_rgb = color_map[color_option]

# 2. ルールの読み込み（古いルールの記憶を強制リセット ttl=1）
@st.cache_data(ttl=1)
def load_rules():
    if os.path.exists('rules.csv'):
        encodings = ['utf-8-sig', 'shift-jis', 'utf-8', 'cp932']
        for enc in encodings:
            try:
                df = pd.read_csv('rules.csv', encoding=enc)
                if '類義語' in df.columns and '統一語句' in df.columns:
                    # 空白や空行を除去
                    df = df.dropna(subset=['類義語', '統一語句'])
                    df['類義語'] = df['類義語'].astype(str).str.strip()
                    df['統一語句'] = df['統一語句'].astype(str).str.strip()
                    df['len'] = df['類義語'].str.len()
                    df = df.sort_values('len', ascending=False)
                    return dict(zip(df['類義語'], df['統一語句']))
            except:
                continue
    return {}

rules_dict = load_rules()

# 3. 修正・熟語保護ロジック (Word/PDF 全ファイル共通)
def apply_rules_to_text(target_text, rules):
    # 【基本保護リスト】
    keep_words = [
        "会員に成長する機会", "会員拡大運動", "会員拡大", "正会員", "新入会員",
        "日本の青年会議所は", "希望をもたらす変革の起点として", 
        "輝く個性が調和する未来を描き", "社会の課題を解決することで", 
        "持続可能な地域を創ることを誓う", "われわれ JAYCEE は", "われわれJAYCEEは",
        "志高き組織ビジョン", "志高き人材育成ビジョン", "志高きまち創造ビジョン"
    ]
    
    # 【追加保護リスト】CSVで「類義語」と「統一語句」が同じものは、保護対象として自動追加
    for k, v in rules.items():
        if k == v and str(k) not in keep_words:
            keep_words.append(str(k))
            
    # 長い言葉から順に保護する（誤作動防止）
    keep_words = sorted(keep_words, key=len, reverse=True)
    
    protected_text = target_text
    placeholders = {}
    
    # 保護対象を一時的に「__SAFE_0001__」のような記号に置き換えて守る
    for i, word in enumerate(keep_words):
        if word in protected_text:
            placeholder = f"__SAFE_{i:04d}__"
            placeholders[placeholder] = word
            protected_text = protected_text.replace(word, placeholder)

    segments = [(protected_text, protected_text, False)]
    
    for wrong, right in rules.items():
        if wrong == right or not wrong: continue
        new_segments = []
        for orig, curr, already_fixed in segments:
            if already_fixed or str(wrong) not in curr:
                new_segments.append((orig, curr, already_fixed))
                continue
            
            parts = curr.split(str(wrong))
            for j, part in enumerate(parts):
                if part != "":
                    new_segments.append((part, part, False))
                if j < len(parts) - 1:
                    new_segments.append((str(wrong), str(right), True))
        segments = new_segments

    # 保護していた言葉を元に戻す
    final_segments = []
    for orig, curr, is_fixed in segments:
        if not is_fixed:
            temp_orig = orig
            for placeholder, original_word in placeholders.items():
                temp_orig = temp_orig.replace(placeholder, original_word)
            final_segments.append((temp_orig, temp_orig, False))
        else:
            final_segments.append((orig, curr, True))
    return final_segments

# --- ファイル修正用関数 ---
def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        parts = apply_rules_to_text(para.text, rules)
        if any(p[2] for p in parts):
            para.text = ""
            for orig, curr, is_fixed in parts:
                run = para.add_run(curr)
                if is_fixed:
                    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
                    run.bold = False
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
                    if any(p[2] for p in parts):
                        cell.value = "".join([p[1] for p in parts])
                        cell.font = Font(color=hex_color, bold=False)
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
                    if any(p[2] for p in parts):
                        paragraph.text = ""
                        for orig, curr, is_fixed in parts:
                            new_run = paragraph.add_run()
                            new_run.text = curr
                            if is_fixed:
                                new_run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
                                new_run.font.bold = False
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# --- PDFチェック用関数 (Wordと全く同じエンジンを使用) ---
def check_pdf(file, rules):
    results = []
    # PDF特有の見えない文字や空白を完全に消去する設定
    invisible_chars = r'[\s\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF\u00A0]+'
    
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if not text: continue
            
            # PDFの文字を「空白のない1本の繋がった文字列」に浄化する
            pure_text = re.sub(invisible_chars, '', text)
            
            # Wordと全く同じ関数に投げて判定させる
            parts = apply_rules_to_text(pure_text, rules)
            
            # 前後の文脈を取得するために、文字列を復元
            full_text = "".join([p[0] for p in parts])
            
            current_idx = 0
            for orig, curr, is_fixed in parts:
                if is_fixed:
                    # NGワードの前後15文字を切り出す
                    start_idx = max(0, current_idx - 15)
                    end_idx = min(len(full_text), current_idx + len(orig) + 15)
                    context = full_text[start_idx:end_idx]
                    
                    results.append({
                        "ページ": i + 1,
                        "NGワード": orig,
                        "修正案": curr,
                        "周辺の文章": f"...{context}..."
                    })
                current_idx += len(orig)
                
    return results

# 4. メイン処理
uploaded_files = st.file_uploader("ファイルをアップロード (Word, Excel, PPT, PDF)", type=["docx", "xlsx", "pptx", "pdf"], accept_multiple_files=True)

if uploaded_files:
    for idx, file in enumerate(uploaded_files):
        ext = file.name.split('.')[-1].lower()
        if ext == "pdf":
            with st.spinner(f"PDF {file.name} を解析中..."):
                pdf_results = check_pdf(file, rules_dict)
                st.subheader(f"📑 {file.name} のチェック結果")
                if pdf_results:
                    st.warning(f"以下の {len(pdf_results)} 箇所で修正が推奨されます。")
                    st.table(pd.DataFrame(pdf_results))
                else:
                    st.success("統一ルールに反する箇所は見つかりませんでした！")
        else:
            with st.spinner(f"{file.name} を処理中..."):
                if ext == "docx": data = repair_docx(file, rules_dict, selected_rgb)
                elif ext == "xlsx": data = repair_xlsx(file, rules_dict, selected_rgb)
                elif ext == "pptx": data = repair_pptx(file, rules_dict, selected_rgb)
                st.download_button(label=f"📥 修正済みの {file.name} を保存", data=data, file_name=f"【修正済】{file.name}", key=f"dl_btn_{idx}")
