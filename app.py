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

st.set_page_config(page_title="川内JC 統一ルール修正ツール", page_icon="⚓", layout="wide")
st.title("⚓ 2026年度 川内JC 統一ルール 修正システム")

# 1. 色の設定
color_option = st.selectbox("修正箇所の文字色", ["赤", "青", "緑", "オレンジ"], index=0)
color_map = {"赤": (255, 0, 0), "青": (0, 0, 255), "緑": (0, 128, 0), "オレンジ": (255, 165, 0)}
selected_rgb = color_map[color_option]

# 2. ルールの読み込み
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        # 文字が長い順に並び替えて、部分一致による誤爆を防ぐ
        df['len'] = df['類義語'].astype(str).str.len()
        df = df.sort_values('len', ascending=False)
        return dict(zip(df['類義語'].astype(str), df['統一語句'].astype(str)))
    return {}

rules_dict = load_rules()

# 3. 修正・色付け・熟語保護ロジック
def apply_rules_to_text(target_text, rules):
    """
    1. 左右が同じルール（保護）は色を変えない
    2. 2文字以下の言葉（会員など）は、前後に漢字がある場合は熟語とみなして無視する
    """
    # (テキスト, 修正確定フラグ) のリスト
    segments = [(target_text, False)]
    
    for wrong, right in rules.items():
        new_segments = []
        for text, already_fixed in segments:
            if already_fixed or wrong not in text:
                new_segments.append((text, already_fixed))
                continue
            
            # 2文字以下の短い言葉（会員など）への境界判定ロジック
            # 前後に漢字 [一-龥] がある場合は、熟語の一部とみなしてスキップ
            if len(wrong) <= 2:
                # 正規表現で「前後に漢字がない場合のみ」を対象に分割
                pattern = f'(?<![一-龥]){re.escape(wrong)}(?![一-龥])'
                split_parts = re.split(f'({pattern})', text)
                
                # re.splitの結果、patternに合致した部分はそのままwrongとして出てくる
                for part in split_parts:
                    if part == wrong:
                        is_real_change = (wrong != right)
                        new_segments.append((right, is_real_change))
                    elif part != "":
                        new_segments.append((part, False))
            else:
                # 3文字以上の長いルールは通常通り分割
                parts = text.split(wrong)
                for i, part in enumerate(parts):
                    if part != "":
                        new_segments.append((part, False))
                    if i < len(parts) - 1:
                        is_real_change = (wrong != right)
                        new_segments.append((right, is_real_change))
        segments = new_segments
        
    return segments

# --- 各種ファイル修正用関数（前回のロジックを継承） ---
def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        if any(str(wrong) in para.text for wrong in rules.keys()):
            parts = apply_rules_to_text(para.text, rules)
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
                    new_val = "".join([p[0] for p in parts])
                    has_change = any(p[1] for p in parts)
                    cell.value = new_val
                    if has_change:
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
                    if any(str(wrong) in combined_text for wrong in rules.keys()):
                        parts = apply_rules_to_text(combined_text, rules)
                        paragraph.text = ""
                        for text, is_fixed in parts:
                            new_run = paragraph.add_run()
                            new_run.text = text
                            if is_fixed:
                                new_run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

uploaded_files = st.file_uploader("ファイルをアップロード", type=["docx", "xlsx", "pptx"], accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        with st.spinner(f"{file.name} を処理中..."):
            if ext == "docx": data = repair_docx(file, rules_dict, selected_rgb)
            elif ext == "xlsx": data = repair_xlsx(file, rules_dict, selected_rgb)
            elif ext == "pptx": data = repair_pptx(file, rules_dict, selected_rgb)
            st.download_button(label=f"📥 修正済みの {file.name} を保存", data=data, file_name=f"【完成版】{file.name}")
