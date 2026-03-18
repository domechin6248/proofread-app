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

# 2. ルールの読み込み（キャッシュを廃止し、常に最新のCSVを読み込む）
def load_rules():
    if os.path.exists('rules.csv'):
        try:
            # utf-8-sigで読み込むことで、Excel保存時のBOM（見えない記号）エラーを防止
            df = pd.read_csv('rules.csv', encoding='utf-8-sig')
        except:
            try:
                df = pd.read_csv('rules.csv', encoding='shift-jis')
            except:
                df = pd.read_csv('rules.csv', encoding='utf-8')
        
        # 列名やデータ内の無駄な空白を削除
        df.columns = df.columns.str.strip()
        if '類義語' in df.columns and '統一語句' in df.columns:
            df = df.dropna(subset=['類義語', '統一語句'])
            df['類義語'] = df['類義語'].astype(str).str.strip()
            df['統一語句'] = df['統一語句'].astype(str).str.strip()
            
            # 長い単語から処理するために文字数でソート
            df['len'] = df['類義語'].str.len()
            df = df.sort_values('len', ascending=False)
            return dict(zip(df['類義語'], df['統一語句']))
    return {}

rules_dict = load_rules()

# 3. 修正・熟語保護ロジック (Word/Excel/PPT用)
def apply_rules_to_text(target_text, rules):
    keep_words = [
        "会員に成長する機会", "会員拡大運動", "正会員", "新入会員",
        "日本の青年会議所は", "希望をもたらす変革の起点として", 
        "輝く個性が調和する未来を描き", "社会の課題を解決することで", 
        "持続可能な地域を創ることを誓う", "われわれ JAYCEE は", "われわれJAYCEEは",
        "志高き組織ビジョン", "志高き人材育成ビジョン", "志高きまち創造ビジョン"
    ]
    
    protected_text = target_text
    placeholders = {}
    for i, word in enumerate(keep_words):
        if word in protected_text:
            placeholder = f"__SAFE_{i}__"
            placeholders[placeholder] = word
            protected_text = protected_text.replace(word, placeholder)

    segments = [(protected_text, False)]
    for wrong, right in rules.items():
        if wrong == right or not wrong: continue
        new_segments = []
        for text, already_fixed in segments:
            if already_fixed or str(wrong) not in text:
                new_segments.append((text, already_fixed))
                continue
            parts = text.split(str(wrong))
            for j, part in enumerate(parts):
                if part != "": new_segments.append((part, False))
                if j < len(parts) - 1: new_segments.append((str(right), True))
        segments = new_segments

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

# --- ファイル修正用関数 ---
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
                    if any(p[1] for p in parts):
                        cell.value = "".join([p[0] for p in parts])
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
                    if any(p[1] for p in parts):
                        paragraph.text = ""
                        for text, is_fixed in parts:
                            new_run = paragraph.add_run()
                            new_run.text = text
                            if is_fixed:
                                new_run.font.color.rgb = PptxRGBColor(rgb[0], rgb[1], rgb[2])
                                new_run.font.bold = False
    out_io = BytesIO()
    prs.save(out_io)
    return out_io.getvalue()

# --- PDFチェック用関数 (新・完全無菌室スキャン版) ---
def check_pdf(file, rules):
    keep_list = [
        "会員に成長する機会", "会員拡大運動", "正会員", "新入会員",
        "日本の青年会議所は", "希望をもたらす変革の起点として", 
        "輝く個性が調和する未来を描き", "社会の課題を解決することで", 
        "持続可能な地域を創ることを誓う", "われわれ JAYCEE は", "われわれJAYCEEは",
        "志高き組織ビジョン", "志高き人材育成ビジョン", "志高きまち創造ビジョン"
    ]
    
    results = []
    # PDF特有の「見えない文字（ゼロ幅スペースなど）」を消去するための正規表現
    invisible_chars = r'[\s\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF\u00A0]+'
    
    with pdfplumber.open(file) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if not text: continue
            
            # PDFの文字を「完全な繋がった文字列」に浄化する
            clean_text = re.sub(invisible_chars, '', text)
            
            for wrong, right in rules.items():
                if wrong == right or not wrong: continue
                clean_wrong = re.sub(invisible_chars, '', str(wrong))
                if not clean_wrong: continue
                
                # 浄化されたテキストの中からNGワードを探す
                for m in re.finditer(re.escape(clean_wrong), clean_text):
                    # 前後20文字を切り出して、保護対象の言葉が含まれていないかチェック
                    start_idx = max(0, m.start() - 20)
                    end_idx = min(len(clean_text), m.end() + 20)
                    context_window = clean_text[start_idx:end_idx]
                    
                    is_protected = False
                    for kw in keep_list:
                        clean_kw = re.sub(invisible_chars, '', str(kw))
                        if not clean_kw: continue
                        
                        # 切り出した文章の中に「会員拡大運動」等が含まれているか
                        kw_start = context_window.find(clean_kw)
                        while kw_start != -1:
                            kw_end = kw_start + len(clean_kw)
                            # 今回見つけたNGワードが、保護対象ワードの「内側」にいるか確認
                            local_m_start = m.start() - start_idx
                            local_m_end = m.end() - start_idx
                            
                            if local_m_start >= kw_start and local_m_end <= kw_end:
                                is_protected = True
                                break
                            
                            kw_start = context_window.find(clean_kw, kw_start + 1)
                        
                        if is_protected:
                            break
                    
                    # 守られていない完全なNGワードのみリストに追加
                    if not is_protected:
                        disp_start = max(0, m.start() - 15)
                        disp_end = min(len(clean_text), m.end() + 15)
                        results.append({
                            "ページ": i + 1,
                            "NGワード": wrong,
                            "修正案": right,
                            "周辺の文章": f"...{clean_text[disp_start:disp_end]}..."
                        })
                        
    # 重複を削除して出力
    if results:
        df = pd.DataFrame(results).drop_duplicates()
        return df.to_dict('records')
    return []

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
