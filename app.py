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

# 1. ページ設定
st.set_page_config(page_title="川内JC 統一ルール修正ツール", page_icon="⚓", layout="wide")

# 2. タイトル
st.title("⚓ 2026年度 川内JC 統一ルール 修正システム")
st.write("1つの文章内に複数の修正箇所があっても、すべて色を変えて修正します。")

# 3. 色の指定
color_option = st.selectbox("修正箇所の文字色", ["赤", "青", "緑", "オレンジ"], index=0)
color_map = {"赤": (255, 0, 0), "青": (0, 0, 255), "緑": (0, 128, 0), "オレンジ": (255, 165, 0)}
selected_rgb = color_map[color_option]

# 4. ルールの読み込み
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        # 文字数の長い順に並び替えて、部分一致による誤爆を防ぐ
        df['len'] = df['類義語'].str.len()
        df = df.sort_values('len', ascending=False)
        return dict(zip(df['類義語'], df['統一語句']))
    return {}

rules_dict = load_rules()

# 5. 修正・色付けロジック
def apply_rules_to_text(target_text, rules):
    """テキスト内のすべてのNGワードを置換し、位置を特定する"""
    changes = []
    temp_text = target_text
    
    # 全ルールをスキャンして置換後のテキストと位置を記録
    for wrong, right in rules.items():
        if str(wrong) in temp_text:
            temp_text = temp_text.replace(str(wrong), f"__FIX__{right}__END__")
    
    # 独自のタグを元に、(テキスト, 色付けフラグ) のリストに分割
    parts = []
    segments = re.split(r'(__FIX__|__END__)', temp_text)
    is_fix = False
    for seg in segments:
        if seg == "__FIX__":
            is_fix = True
        elif seg == "__END__":
            is_fix = False
        elif seg != "":
            parts.append((seg, is_fix))
    return parts

def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        if any(str(wrong) in para.text for wrong in rules.keys()):
            original_text = para.text
            parts = apply_rules_to_text(original_text, rules)
            para.text = "" # 段落をクリアして再構築
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
                    new_val = cell.value
                    changed = False
                    for wrong, right in rules.items():
                        if str(wrong) in new_val:
                            new_val = new_val.replace(str(wrong), str(right))
                            changed = True
                    if changed:
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
                        paragraph.text = "" # 一旦クリア
                        for text, color_flag in parts:
                            new_run = paragraph.add_run()
                            new_run.text = text
                            if color_flag:
                                new_run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# 6. メイン画面
uploaded_files = st.file_uploader("ファイルをアップロード", type=["docx", "xlsx", "pptx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        with st.spinner(f"{file.name} を処理中..."):
            if ext == "docx":
                data = repair_docx(file, rules_dict, selected_rgb)
            elif ext == "xlsx":
                data = repair_xlsx(file, rules_dict, selected_rgb)
            elif ext == "pptx":
                data = repair_pptx(file, rules_dict, selected_rgb)
            
            st.download_button(label=f"📥 修正済みの {file.name} を保存", data=data, file_name=f"【修正済】{file.name}")
