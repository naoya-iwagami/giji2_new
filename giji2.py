import os  
import openai  
import docx  
import streamlit as st  
import tempfile  

# ページの設定をwideにする  
st.set_page_config(layout="wide")  

# 環境変数からAPIキーを取得  
api_key = os.getenv("OPENAI_API_KEY")  
if api_key is None:  
    st.error("APIキーが設定されていません。環境変数OPENAI_API_KEYを設定してください。")  
    st.stop()  

# OpenAI APIの設定  
openai.api_type = "azure"  
openai.api_base = "https://test-chatgpt-pm-1.openai.azure.com/"  
openai.api_version = "2023-03-15-preview"  
openai.api_key = api_key

# スクリプトのディレクトリを取得
script_dir = os.path.dirname(os.path.abspath(__file__))

# 辞書ファイルのパスをプログラム内に記載
dictionary_filename = "AI教育用専門用語集20240730.txt"
dictionary_path = os.path.join(script_dir, dictionary_filename)

def read_docx(file_path):  
    try:  
        doc = docx.Document(file_path)  
        full_text = []  
        for para in doc.paragraphs:  
            full_text.append(para.text)  
        return '\n'.join(full_text)  
    except Exception as e:  
        st.error(f"Error reading docx file: {e}")  
        return None

def read_dictionary(dictionary_path):
    if os.path.exists(dictionary_path):
        # 辞書ファイルの内容を文字列として読み込む
        try:
            with open(dictionary_path, 'r', encoding='utf-8') as f:
                dictionary_content = f.read()
            return dictionary_content
        except Exception as e:
            st.error(f"辞書ファイルの読み込み中にエラーが発生しました: {e}")
            return ""
    else:
        st.error("指定された辞書ファイルが存在しません。")
        return ""

def create_summary(full_text, include_names, dictionary_content):  
    try:  
        system_message = (
            "あなたはプロフェッショナルな議事録作成者です。"
            "以下の専門用語集を参考に、全文を基にできるだけ詳細かつ具体的な議事録を作成してください。"
            "専門用語集には、このドメインにおける重要な用語とその解説が含まれています。"
            "議事録には以下の3つの項目のみを含めてください:\n"
            "1. 会議概要\n"
            "2. 詳細な議論内容（具体的な発言や数値データを含む）\n"
            "3. アクションアイテム\n\n"
            "これら以外の項目は含めないでください。"
            "重要なポイントや具体例を盛り込み、できるだけ箇条書きではなく文章で深みのある議事録を作成してください。"
        )

        if include_names:  
            system_message += "\n参加者の名前や発言者が特定できる場合はそれも含めてください。"  
        else:  
            system_message += "\n参加者の名前や発言者の情報は含めないでください。"

        user_message = f"専門用語集:\n{dictionary_content}\n\n全文:\n{full_text}"

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": user_message}  
            ],  
            temperature=0,
            max_tokens=8000
        )  
        summary = response["choices"][0]["message"]["content"]
        return summary
    except Exception as e:  
        st.error(f"Unexpected error in create_summary: {e}")  
        return full_text

def hiroyuki_comments(summary):  
    try:  
        system_message = "あなたは論破王ひろゆきです。以下の議事録について皮肉や論理的な切り口で、少し挑発的かつ軽い皮肉を交えて指摘を行ってください。また、話の流れが少しおかしい部分や矛盾がありそうな部分を鋭く指摘し、相手を諭すようなトーンでコメントしてください。"  

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は議事録です。各項目に指摘を入れてください。\n\n{summary}"}  
            ],  
            temperature=0.1,
            max_tokens=2000
        )  
        return response["choices"][0]["message"]["content"]  
    except Exception as e:  
        st.error(f"Unexpected error in hiroyuki_comments: {e}")  

def consulting_advice(summary):  
    try:  
        system_message = "あなたはプロフェッショナルなコンサルタントです。以下の議事録を基に、ビジネスの改善点や次のステップについて具体的なアドバイスをしてください。"

        response = openai.ChatCompletion.create(  
            engine="pm-GPT4o-mini",  
            messages=[  
                {"role": "system", "content": system_message},  
                {"role": "user", "content": f"以下は議事録です。これを基にアドバイスをしてください。\n\n{summary}"}  
            ],  
            temperature=0.3,
            max_tokens=2000
        )  
        return response["choices"][0]["message"]["content"]  
    except Exception as e:  
        st.error(f"Unexpected error in consulting_advice: {e}")  

# 辞書ファイルの内容を読み込み
dictionary_content = read_dictionary(dictionary_path)

# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  

# 参加者の名前や発言者情報を含めるかどうかのチェックボックス  
include_names = st.checkbox("参加者の名前や発言者の情報を含める", value=True)  

# ひろゆきの指摘を受けるかどうかのチェックボックス  
with_hiroyuki = st.checkbox("ひろゆきの指摘を受ける", value=False)  

# コンサルティングのアドバイスを受けるかどうかのチェックボックス  
with_consulting = st.checkbox("コンサルティングのアドバイスを受ける", value=False)  

# 複数ファイルのアップロード
uploaded_files = st.file_uploader("ファイルをアップロード (最大3つ)", type="docx", accept_multiple_files=True)  

if uploaded_files:  
    if len(uploaded_files) > 3:  
        st.error("最大3つのファイルまでアップロードできます。")  
    else:  
        # 複数のファイルの内容を結合  
        all_texts = ""  
        for uploaded_file in uploaded_files:  
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:  
                tmp_file.write(uploaded_file.read())  
                tmp_file_path = tmp_file.name  
            text = read_docx(tmp_file_path)  
            if text:  
                all_texts += text + "\n"  

        if all_texts:  
            # 議事録を作成  
            summary = create_summary(all_texts, include_names, dictionary_content)

            # レイアウトの決定
            if with_hiroyuki and with_consulting:  
                col1, col2, col3 = st.columns(3)  
            elif with_hiroyuki or with_consulting:  
                col1, col2 = st.columns(2)  
            else:  
                col1 = st.container()  

            with col1:  
                st.write(summary)

            if with_hiroyuki:  
                with col2 if 'col2' in locals() else col1:  
                    st.header("ひろゆきの指摘")  
                    comments = hiroyuki_comments(summary)  
                    if comments:  
                        st.write(comments)  
                    else:  
                        st.error("ひろゆきの指摘の生成に失敗しました。")  

            if with_consulting:  
                with col3 if 'col3' in locals() else col2:  
                    st.header("コンサルティングのアドバイス")  
                    advice = consulting_advice(summary)  
                    if advice:  
                        st.write(advice)  
                    else:  
                        st.error("コンサルティングのアドバイスの生成に失敗しました。")  

        else:  
            st.error("ファイルの読み込みに失敗しました。")  
