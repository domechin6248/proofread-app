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

# 色の設定
color_option = st.selectbox("修正箇所の文字色", ["赤", "青", "緑", "オレンジ"], index=0)
color_map = {"赤": (255, 0, 0), "青": (0, 0, 255), "緑": (0, 128, 0), "オレンジ": (255, 165, 0)}
selected_rgb = color_map[color_option]

@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        # 文字が長い順に処理することで誤爆を防ぐ
        df['len'] = df['類義語'].str.len()
        df = df.sort_values('len', ascending=False)
        return dict(zip(df['類義語'], df['統一語句']))
    return {}

rules_dict = load_rules()

def apply_rules_to_text(target_text, rules):
    """
    正規表現を使い、単語の境界（他の熟語の一部ではない場合）を意識して置換する
    """
    temp_text = target_text
    # 修正が必要な言葉を「安全な一時タグ」に置き換える
    for wrong, right in rules.items():
        if str(wrong) in temp_text:
            # 特殊ルール：2文字以下の短い単語（会員など）は、
            # 前後に他の漢字や文字がくっついていない場合のみ置換を試みる（簡易的な境界判定）
            if len(str(wrong)) <= 2:
                # 前後に文字がない、または特定の記号がある場合のみ反応させる
                # 熟語（会員拡大など）を避けるための設定
                pattern = r'(?<![一-龥])' + re.escape(str(wrong)) + r'(?![一-龥])'
                temp_text = re.sub(pattern, f"__FIX__{right}__END__", temp_text)
            else:
                temp_text = temp_text.replace(str(wrong), f"__FIX__{right}__END__")
    
    parts = []
    segments = re.split(r'(__FIX__|__END__)', temp_text)
    is_fix = False
    for seg in segments:
        if seg == "__FIX__": is_fix = True
        elif seg == "__END__": is_fix = False
        elif seg != "": parts.append((seg, is_fix))
    return parts

# --- 各種ファイル修正用関数（中身は前回の「複数箇所対応版」と同様） ---
def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        if any(str(wrong) in para.text for wrong in rules.keys()):
            parts = apply_rules_to_text(para.text, rules)
            para.text = ""
            for text, color_flag in parts:
                run = para.add_run(text)
                if color_flag:
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
                    if new_val != cell.value:
                        cell.value = new_val
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
                        for text, color_flag in parts:
                            new_run = paragraph.add_run()
                            new_run.text = text
                            if color_flag:
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
            st.download_button(label=f"📥 修正済みの {file.name} を保存", data=data, file_name=f"【修正済】{file.name}")
