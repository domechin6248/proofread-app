import streamlit as st
import pandas as pd
import docx
import os

st.set_page_config(page_title="2026年度 校正システム", layout="wide")

st.title("📋 2026年度 統一ルール校正ツール")
st.write("ファイルをアップロードすると、自動で誤字脱字や書式をチェックします。")

# 統一語句データの読み込み
@st.cache_data
def load_rules():
    # 提供いただいたCSVファイルを読み込む
    df = pd.read_csv('rules.csv')
    return dict(zip(df['類義語'], df['統一語句']))

try:
    rules_dict = load_rules()
except:
    st.error("ルール設定ファイル(rules.csv)が見つかりません。")
    rules_dict = {}

uploaded_files = st.file_uploader("Wordファイルをアップロード", type="docx", accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"🔍 {uploaded_file.name} のチェック結果")
        doc = docx.Document(uploaded_file)
        results = []

        for i, para in enumerate(doc.paragraphs):
            # 1. 統一語句チェック
            for wrong, right in rules_dict.items():
                if str(wrong) in para.text:
                    results.append({"行": i+1, "種別": "語句", "内容": f"「{wrong}」→「{right}」に修正してください"})
            
            # 2. 字体チェック (MS 明朝)
            for run in para.runs:
                if run.font.name and 'MS' not in run.font.name and '明朝' not in run.font.name:
                    results.append({"行": i+1, "種別": "書式", "内容": f"フォントを確認してください: {run.font.name}"})

        if results:
            st.table(pd.DataFrame(results))
        else:
            st.success("問題ありません！")
