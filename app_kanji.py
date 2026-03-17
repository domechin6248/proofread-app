import streamlit as st
import docx
import google.generativeai as genai
import os

# 1. ページ設定
st.set_page_config(page_title="川内JC 監事審査AI", page_icon="⚖️", layout="wide")
st.title("⚖️ 2026年度 川内JC 監事専用・事業計画分析システム")
st.write("AIが計画書を読み込み、論理の整合性と改善案を提示します。")

# 2. APIキーの設定（サイドバー）
with st.sidebar:
    st.header("🔑 設定")
    api_key = st.text_input("Gemini API Keyを入力してください", type="password")
    if api_key:
        genai.configure(api_key=api_key)

# 3. 分析実行関数
def analyze_with_ai(text):
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # 監事の視点（プロンプト）
    prompt = f"""
    あなたは川内青年会議所（JC）の「監事」です。36歳の不動産業経営者という背景を持ち、
    実務に厳しく、論理の整合性を最も重視します。
    以下の事業計画書のテキストを読み、監事講評を作成してください。

    【審査のポイント】
    1. 事業背景：地域の課題（薩摩川内市）が具体的か？データに基づいているか？
    2. 事業目的：背景の課題を解決する「一本筋」が通ったゴールか？
    3. 事業内容：目的達成のために最短距離の手法か？メンバーの負担は適正か？
    4. 検証方法：LOMの会費を使う価値を数値で証明できるか？

    【出力形式】
    ・各項目ごとの「現状分析」と「具体的アドバイス」
    ・全体の「一本筋」が通っているかの判定（合格・要再考）
    ・監事としての厳しいが愛のある総評

    【計画書テキスト】
    {text}
    """
    response = model.generate_content(prompt)
    return response.text

# 4. Wordファイル読み込み
def get_text_from_docx(file):
    doc = docx.Document(file)
    return "\\n".join([p.text for p in doc.paragraphs if p.text.strip()])

# 5. メイン画面
uploaded_file = st.file_uploader("審査する計画書(Word)をドロップしてください", type=["docx"])

if uploaded_file:
    if not api_key:
        st.warning("左側のサイドバーにAPIキーを入力してください。")
    else:
        text_content = get_text_from_docx(uploaded_file)
        if st.button("監事審査を開始する"):
            with st.spinner("監事が内容を精査中..."):
                try:
                    result = analyze_with_ai(text_content)
                    st.markdown("---")
                    st.header("📢 監事審査結果")
                    st.markdown(result)
                except Exception as e:
                    st.error(f"エラーが発生しました: {{e}}")
