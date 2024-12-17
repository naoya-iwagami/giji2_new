import os  
import openai  
import docx  
import streamlit as st  
import tiktoken  # トークン数計算のために追加  
  
# ページの設定をwideにする  
st.set_page_config(layout="wide")  
  
# 環境変数からAPIキーを取得  
api_key = os.getenv("OPENAI_API_KEY")  
if not api_key:  
    st.error("APIキーが設定されていません。環境変数OPENAI_API_KEYを設定してください。")  
    st.stop()  
  
# OpenAI APIの設定  
openai.api_type = "azure"  
openai.api_base = "https://test-chatgpt-pm-1.openai.azure.com/"  # ここはご自身のエンドポイントに合わせて変更してください  
openai.api_version = "2024-10-01-preview"  # エンジンに合わせてAPIバージョンを指定  
openai.api_key = api_key  
  
# エンコーディングの設定（明示的に 'cl100k_base' を指定）  
encoding = tiktoken.get_encoding("cl100k_base")  # 'o1-preview' エンジンに適したエンコーディングを指定  
  
def read_docx(file):  
    try:  
        doc = docx.Document(file)  
        full_text = [para.text for para in doc.paragraphs]  
        return '\n'.join(full_text)  
    except Exception as e:  
        st.error(f"Error reading docx file: {e}")  
        return None  
  
def num_tokens_from_messages(messages, encoding):  
    """メッセージからトークン数を計算"""  
    num_tokens = 0  
    for message in messages:  
        num_tokens += 4  # メッセージのヘッダ部分のトークン数  
        for key, value in message.items():  
            num_tokens += len(encoding.encode(value))  
    num_tokens += 2  # 会話の終端部分のトークン数  
    return num_tokens  
  
def split_text_into_chunks(text, max_tokens_per_chunk, encoding):  
    """テキストを指定したトークン数以下のチャンクに分割"""  
    tokens = encoding.encode(text)  
    chunks = []  
    start = 0  
    while start < len(tokens):  
        end = start + max_tokens_per_chunk  
        chunk_tokens = tokens[start:end]  
        chunk_text = encoding.decode(chunk_tokens)  
        chunks.append(chunk_text)  
        start = end  
    return chunks  
  
def create_summary(full_text, include_names, include_time, desired_char_count):  
    try:  
        # モデルの最大コンテキスト長（入力＋出力トークン数の合計）  
        max_context_length = 160000  
  
        # モデルの最大出力トークン数  
        max_output_tokens = 32768  
  
        # チャンクごとの最大入力トークン数を設定  
        # 入力トークン数 + 出力トークン数 <= max_context_length となるようにする  
        buffer_tokens = 5000  # システムメッセージや余裕のためのトークン数  
        max_chunk_tokens = max_context_length - max_output_tokens - buffer_tokens  
  
        # テキストをチャンクに分割  
        chunks = split_text_into_chunks(full_text, max_chunk_tokens, encoding)  
  
        # 各チャンクの要約を格納するリスト  
        chunk_summaries = []  
  
        # 各チャンクごとに要約を生成  
        for i, chunk in enumerate(chunks):  
            st.write(f"チャンク {i+1}/{len(chunks)} を処理中...")  
  
            # システムメッセージの内容を用意  
            system_message = (  
                "あなたはプロフェッショナルな議事録作成者です。"  
                "与えられたテキストを基にできるだけ詳細かつ具体的な議事録を作成してください。"  
                "議事録には以下の項目を含めてください:\n"  
                "1. 会議開催日"  
            )  
            if include_time:  
                system_message += "\n2. 会議開催時間"  
                content_points = "\n3. 会議概要\n4. 詳細な議論内容（具体的な発言や数値データを含む）\n5. アクションアイテム\n"  
            else:  
                content_points = "\n2. 会議概要\n3. 詳細な議論内容（具体的な発言や数値データを含む）\n4. アクションアイテム\n"  
            system_message += content_points  
            system_message += (  
                "これら以外の項目は含めないでください。"  
                "重要なポイントや具体例を盛り込み、できるだけ箇条書きではなく文章で深みのある議事録を作成してください。"  
                f"\n\n議事録の文字数はおおよそ{desired_char_count}文字にしてください。"  
                "詳細かつ包括的な内容を提供し、指定された文字数に近づけてください。"  
                "情報が不足している場合は、推測や架空の情報を追加しないでください。"  
            )  
            if include_names:  
                system_message += "\n参加者の名前や発言者が特定できる場合はそれも含めてください。"  
            else:  
                system_message += "\n参加者の名前や発言者の情報は含めないでください。"  
  
            # ユーザーメッセージを作成（システムメッセージと結合）  
            user_message = f"{system_message}\n\n以下はテキストの一部です。\n\n{chunk}"  
  
            messages = [  
                {"role": "user", "content": user_message}  
            ]  
  
            # メッセージのトークン数を計算  
            prompt_tokens = num_tokens_from_messages(messages, encoding)  
  
            # 最大出力トークン数（モデルの制限内で、指定された最大値）  
            max_completion_tokens = max_output_tokens  
  
            # 総トークン数がモデルの最大コンテキスト長を超えないように調整  
            if prompt_tokens + max_completion_tokens > max_context_length:  
                max_completion_tokens = max_context_length - prompt_tokens  
  
            if max_completion_tokens <= 0:  
                st.error("チャンクが長すぎます。チャンクサイズを小さくしてください。")  
                return None  
  
            response = openai.ChatCompletion.create(  
                engine='o1-preview',  
                messages=messages,  
                max_completion_tokens=max_completion_tokens  
            )  
  
            chunk_summary = response["choices"][0]["message"]["content"]  
            chunk_summaries.append(chunk_summary)  
  
        # すべてのチャンクの要約を結合  
        combined_summary = "\n".join(chunk_summaries)  
  
        # 最終的な要約をさらに要約（オプション）  
        st.write("すべてのチャンクの要約を統合しています...")  
  
        # 最終要約時のメッセージを作成  
        final_system_message = (  
            "あなたはプロフェッショナルな議事録作成者です。"  
            "以下の要約を統合して、全体の議事録を作成してください。"  
            f"\n\n議事録の文字数はおおよそ{desired_char_count}文字にしてください。"  
            "重要なポイントを含め、簡潔かつ明確な議事録を作成してください。"  
        )  
  
        final_user_message = f"{final_system_message}\n\n以下は各チャンクの要約です。\n\n{combined_summary}"  
  
        final_messages = [  
            {"role": "user", "content": final_user_message}  
        ]  
  
        prompt_tokens = num_tokens_from_messages(final_messages, encoding)  
  
        # 最大出力トークン数（モデルの制限内で、指定された最大値）  
        max_completion_tokens = max_output_tokens  
  
        # 総トークン数がモデルの最大コンテキスト長を超えないように調整  
        if prompt_tokens + max_completion_tokens > max_context_length:  
            max_completion_tokens = max_context_length - prompt_tokens  
  
        if max_completion_tokens <= 0:  
            st.error("最終要約が長すぎます。チャンクサイズを小さくしてください。")  
            return None  
  
        final_response = openai.ChatCompletion.create(  
            engine='o1-preview',  
            messages=final_messages,  
            max_completion_tokens=max_completion_tokens  
        )  
  
        final_summary = final_response["choices"][0]["message"]["content"]  
        return final_summary  
  
    except Exception as e:  
        st.error(f"Unexpected error in create_summary: {e}")  
        return None  
  
def hiroyuki_comments(summary):  
    try:  
        # システムメッセージの内容を用意  
        system_message = "あなたは論破王ひろゆきです。以下の議事録について、少し挑発的かつ軽い皮肉を交えて指摘を行ってください。"  
  
        # ユーザーメッセージを作成（システムメッセージと結合）  
        user_message = f"{system_message}\n\n以下は議事録です。各項目に指摘を入れてください。\n\n{summary}"  
  
        messages = [  
            {"role": "user", "content": user_message}  
        ]  
  
        # メッセージのトークン数を計算  
        prompt_tokens = num_tokens_from_messages(messages, encoding)  
  
        max_context_length = 160000  
        max_output_tokens = 32768  
  
        # 最大出力トークン数（モデルの制限内で、指定された最大値）  
        max_completion_tokens = max_output_tokens  
  
        # 総トークン数がモデルの最大コンテキスト長を超えないように調整  
        if prompt_tokens + max_completion_tokens > max_context_length:  
            max_completion_tokens = max_context_length - prompt_tokens  
  
        if max_completion_tokens <= 0:  
            st.error("議事録が長すぎます。議事録の文字数を減らしてください。")  
            return None  
  
        response = openai.ChatCompletion.create(  
            engine='o1-preview',  
            messages=messages,  
            max_completion_tokens=max_completion_tokens  
        )  
        return response["choices"][0]["message"]["content"]  
    except Exception as e:  
        st.error(f"Unexpected error in hiroyuki_comments: {e}")  
        return None  
  
def consulting_advice(summary):  
    try:  
        # システムメッセージの内容を用意  
        system_message = "あなたはプロフェッショナルなコンサルタントです。以下の議事録を基に、ビジネスの改善点や次のステップについて具体的なアドバイスをしてください。"  
  
        # ユーザーメッセージを作成（システムメッセージと結合）  
        user_message = f"{system_message}\n\n以下は議事録です。これを基にアドバイスをしてください。\n\n{summary}"  
  
        messages = [  
            {"role": "user", "content": user_message}  
        ]  
  
        # メッセージのトークン数を計算  
        prompt_tokens = num_tokens_from_messages(messages, encoding)  
  
        max_context_length = 160000  
        max_output_tokens = 32768  
  
        # 最大出力トークン数（モデルの制限内で、指定された最大値）  
        max_completion_tokens = max_output_tokens  
  
        # 総トークン数がモデルの最大コンテキスト長を超えないように調整  
        if prompt_tokens + max_completion_tokens > max_context_length:  
            max_completion_tokens = max_context_length - prompt_tokens  
  
        if max_completion_tokens <= 0:  
            st.error("議事録が長すぎます。議事録の文字数を減らしてください。")  
            return None  
  
        response = openai.ChatCompletion.create(  
            engine='o1-preview',  
            messages=messages,  
            max_completion_tokens=max_completion_tokens  
        )  
        return response["choices"][0]["message"]["content"]  
    except Exception as e:  
        st.error(f"Unexpected error in consulting_advice: {e}")  
        return None  
  
# StreamlitのUI設定  
st.title("議事録アプリ")  
st.write("docxファイルをアップロードして、その内容から議事録を作成します。")  
  
# 参加者の名前や発言者情報を含めるかどうかのチェックボックス  
include_names = st.checkbox("参加者の名前や発言者の情報を含める", value=False)  
  
# 会議の開催時間を含めるかどうかのチェックボックス（デフォルトは含めない）  
include_time = st.checkbox("会議の開催時間を含める", value=False)  
  
# ひろゆきの指摘を受けるかどうかのチェックボックス  
with_hiroyuki = st.checkbox("ひろゆきの指摘を受ける", value=False)  
  
# コンサルティングのアドバイスを受けるかどうかのチェックボックス  
with_consulting = st.checkbox("コンサルティングのアドバイスを受ける", value=False)  
  
# 議事録の文字数を指定する入力フィールド（最大値を30,000に設定）  
desired_char_count = st.number_input(  
    "議事録の文字数を指定してください（例：500～30000文字）",  
    min_value=500, max_value=30000, value=4000, step=100  
)  
  
# 複数ファイルのアップロード  
uploaded_files = st.file_uploader(  
    "ファイルをアップロード (最大3つ)",  
    type="docx",  
    accept_multiple_files=True  
)  
  
if uploaded_files:  
    if len(uploaded_files) > 3:  
        st.error("最大3つのファイルまでアップロードできます。")  
    else:  
        total_char_count = 0  
        file_char_counts = []  
        all_texts = ""  # 全文テキストを保存する変数  
  
        for uploaded_file in uploaded_files:  
            text = read_docx(uploaded_file)  
            if text:  
                char_count = len(text)  
                total_char_count += char_count  
                file_char_counts.append((uploaded_file.name, char_count))  
                all_texts += text + "\n"  
            else:  
                st.error(f"{uploaded_file.name} の読み込みに失敗しました。")  
  
        # 各ファイルの文字数と合計文字数を表示  
        st.write("アップロードされたファイルの文字数:")  
        for file_name, char_count in file_char_counts:  
            st.write(f"- {file_name}: {char_count}文字")  
        st.write(f"**合計文字数: {total_char_count}文字**")  
  
        # 議事録作成のボタンを配置  
        if st.button('議事録を作成'):  
            if all_texts:  
                # 議事録を作成  
                summary = create_summary(  
                    all_texts,  
                    include_names,  
                    include_time,  
                    desired_char_count  
                )  
  
                if summary is None:  
                    st.error("議事録の生成に失敗しました。")  
                    st.stop()  
  
                # レイアウトの決定  
                if with_hiroyuki and with_consulting:  
                    col1, col2, col3 = st.columns(3)  
                elif with_hiroyuki or with_consulting:  
                    col1, col2 = st.columns(2)  
                else:  
                    col1 = st.container()  
  
                with col1:  
                    st.header("作成された議事録")  
                    st.write(summary)  
  
                if with_hiroyuki:  
                    comments = hiroyuki_comments(summary)  
                    if comments:  
                        if 'col2' in locals():  
                            with col2:  
                                st.header("ひろゆきの指摘")  
                                st.write(comments)  
                        else:  
                            st.header("ひろゆきの指摘")  
                            st.write(comments)  
                    else:  
                        st.error("ひろゆきの指摘の生成に失敗しました。")  
  
                if with_consulting:  
                    advice = consulting_advice(summary)  
                    if advice:  
                        if 'col3' in locals():  
                            with col3:  
                                st.header("コンサルティングのアドバイス")  
                                st.write(advice)  
                        elif 'col2' in locals():  
                            with col2:  
                                st.header("コンサルティングのアドバイス")  
                                st.write(advice)  
                    else:  
                        st.error("コンサルティングのアドバイスの生成に失敗しました。")  
  
            else:  
                st.error("ファイルの読み込みに失敗しました。")  