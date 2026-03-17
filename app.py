import streamlit as st
import pandas as pd
import docx
from docx.shared import RGBColor
import openpyxl
from openpyxl.styles import Font
from pptx import Presentation
from pptx.dml.color import RGBColor as PptxRGBColor
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
st.write("修正箇所の文字色を変更して、自動修正済みのファイルを生成します。")

# 3. 色の指定（ここで色を選べるようにしました）
color_option = st.selectbox(
    "修正箇所の文字色を選んでください",
    ["赤", "青", "緑", "オレンジ"],
    index=0
)

# 色の定義
color_map = {
    "赤": (255, 0, 0),
    "青": (0, 0, 255),
    "緑": (0, 128, 0),
    "オレンジ": (255, 165, 0)
}
selected_rgb = color_map[color_option]

# 4. ルールの読み込み
@st.cache_data
def load_rules():
    if os.path.exists('rules.csv'):
        df = pd.read_csv('rules.csv')
        return dict(zip(df['類義語'], df['統一語句']))
    else:
        st.error("rules.csv が見つかりません。")
        return {}

rules_dict = load_rules()

# 5. 各形式の修正処理（色付け機能付き）
def repair_docx(file, rules, rgb):
    doc = docx.Document(file)
    for para in doc.paragraphs:
        for wrong, right in rules.items():
            if str(wrong) in para.text:
                # 既存のテキストを分割して置換箇所のみ色を変える
                original_text = para.text
                if str(wrong) in original_text:
                    para.text = "" # 一旦クリア
                    parts = original_text.split(str(wrong))
                    for i, part in enumerate(parts):
                        para.add_run(part)
                        if i < len(parts) - 1:
                            new_run = para.add_run(str(right))
                            new_run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
                            new_run.bold = True # 目立たせるために太字
    out_io = BytesIO()
    doc.save(out_io)
    return out_io.getvalue()

def repair_xlsx(file, rules, rgb):
    wb = openpyxl.load_workbook(file)
    target_font = Font(color=f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}", bold=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for wrong, right in rules.items():
                        if str(wrong) in cell.value:
                            cell.value = cell.value.replace(str(wrong), str(right))
                            cell.font = target_font # セル全体の文字色が変わります
    out_io = BytesIO()
    wb.save(out_io)
    return out_io.getvalue()

def repair_pptx(file, rules, rgb):
    prs = Presentation(file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for wrong, right in rules.items():
                        if str(wrong) in paragraph.text:
                            # パワポは構造が複雑なため全体置換＋色付け
                            for run in paragraph.runs:
                                if str(wrong) in run.text:
                                    run.text = run.text.replace(str(wrong), str(right))
                                    run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# 6. メイン処理
uploaded_files = st.file_uploader(
    "ファイルをアップロードしてください", 
    type=["docx", "xlsx", "pptx"], 
    accept_multiple_files=True
)

if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        with st.status(f"🛠 {file.name} を修正・色付け中...", expanded=False):
            if ext == "docx":
                repaired_data = repair_docx(file, rules_dict, selected_rgb)
                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == "xlsx":
                repaired_data = repair_xlsx(file, rules_dict, selected_rgb)
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif ext == "pptx":
                repaired_data = repair_pptx(file, rules_dict, selected_rgb)
                mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            
            st.download_button(
                label=f"📥 修正済みの {file.name} を保存",
                data=repaired_data,
                file_name=f"【色付修正】{file.name}",
                mime=mime,
                key=file.name
            )
