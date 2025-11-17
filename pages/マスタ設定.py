import streamlit as st
import pandas as pd
import os
import glob
import io

# st.set_page_config は削除 (streamlit_app.py で設定済みのため)

# --- ページ表示 ---

st.markdown("""
    <style>
        [data-testid="stSidebarNav"] ul { display: none; }
        .custom-title {
            font-size: 2.1rem; font-weight: 600; color: #3A322E;
            padding-bottom: 10px; border-bottom: 3px solid #FF9933; margin-bottom: 25px;
        }
        .stApp { background: #fff5e6; }
    </style>
""", unsafe_allow_html=True)

# --- ▼▼ ここからサイドバー ▼▼ ---
# すべてのページで同じメニューを表示
st.sidebar.title("メニュー")
st.sidebar.page_link("streamlit_app.py", label="数出表 変換", icon="📄")
st.sidebar.page_link("pages/シール.py", label="シール貼付 変換", icon="🏷️")
st.sidebar.page_link("pages/マスタ設定.py", label="マスタ設定", icon="⚙️")
show_debug = st.sidebar.checkbox("デバッグ情報を表示", value=False)
# --- ▲▲ ここまでサイドバー ▲▲ ---


st.markdown('<p class="custom-title">マスタデータ設定</p>', unsafe_allow_html=True)

st.write("更新するマスタの確認、および新しいCSVファイルのアップロードができます。")

# --- マスタ読み込み/保存/再読み込みのための関数 ---
# (streamlit_app.py と同じ関数を、このページでも使えるように定義します)
def load_master_data(file_prefix, default_columns):
    """
    指定されたプレフィックスに一致する最新のCSVファイルからマスタデータを読み込む
    """
    list_of_files = glob.glob(os.path.join('.', f'{file_prefix}*.csv'))
    if not list_of_files:
        if show_debug:
            st.warning(f"'{file_prefix}*.csv' が見つかりません。空のDataFrameを返します。")
        return pd.DataFrame(columns=default_columns)
    
    latest_file = max(list_of_files, key=os.path.getmtime)
    if show_debug:
        st.info(f"読み込み中: {latest_file}")
        
    encodings = ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']
    for encoding in encodings:
        try:
            df = pd.read_csv(latest_file, encoding=encoding, dtype=str).fillna('')
            if not df.empty: 
                return df
        except Exception as e:
            if show_debug:
                st.warning(f"{latest_file} の読み込み失敗 (encoding: {encoding}): {e}")
            continue
            
    st.error(f"'{latest_file}' の読み込みに失敗しました。")
    return pd.DataFrame(columns=default_columns)

def save_uploaded_csv(uploaded_file, file_prefix):
    """
    アップロードされたCSVファイルを保存する。
    古い同じプレフィックスのファイルは削除する。
    """
    try:
        # 古いファイルを削除
        old_files = glob.glob(os.path.join('.', f'{file_prefix}*.csv'))
        for f in old_files:
            os.remove(f)
            if show_debug:
                st.info(f"古いファイルを削除しました: {f}")
                
        # 新しいファイルを保存
        file_path = os.path.join('.', uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        st.success(f"'{uploaded_file.name}' を保存しました。")
        return True
    except Exception as e:
        st.error(f"ファイルの保存中にエラーが発生しました: {e}")
        return False

def reload_all_masters():
    """
    セッション状態のマスタデータを強制的に再読み込みする
    """
    st.session_state.master_df = load_master_data("商品マスタ一覧", ['商品予定名', 'パン箱入数', '商品名', '売価単価', '弁当区分'])
    st.session_state.customer_master_df = load_master_data("得意先マスタ一覧", ['得意先ＣＤ', '得意先名'])
    st.info("マスタデータを再読み込みしました。")

# --- メインロジック ---

# 1. プレビュー表示
st.subheader("現在の商品マスタデータ（プレビュー）")
if 'master_df' not in st.session_state or st.session_state.master_df.empty:
    st.warning("商品マスタが読み込まれていません。")
else:
    st.dataframe(st.session_state.master_df.head(10)) # 先頭10件のみ表示

st.subheader("現在の得意先マスタデータ（プレビュー）")
if 'customer_master_df' not in st.session_state or st.session_state.customer_master_df.empty:
    st.warning("得意先マスタが読み込まれていません。")
else:
    st.dataframe(st.session_state.customer_master_df.head(10)) # 先頭10件のみ表示

st.divider()

# 2. アップロードセクション
st.subheader("マスタファイルのアップロード")

# 商品マスタのアップロード
uploaded_product_master = st.file_uploader(
    "新しい「商品マスタ一覧...csv」をアップロード",
    type="csv",
    key="product_uploader"
)
if uploaded_product_master is not None:
    if "商品マスタ一覧" in uploaded_product_master.name:
        if save_uploaded_csv(uploaded_product_master, "商品マスタ一覧"):
            reload_all_masters() # 保存成功したら再読み込み
            st.rerun() # ページをリフレッシュしてプレビューを更新
    else:
        st.error("ファイル名に「商品マスタ一覧」が含まれていません。")

# 得意先マスタのアップロード
uploaded_customer_master = st.file_uploader(
    "新しい「得意先マスタ一覧...csv」をアップロード",
    type="csv",
    key="customer_uploader"
)
if uploaded_customer_master is not None:
    if "得意先マスタ一覧" in uploaded_customer_master.name:
        if save_uploaded_csv(uploaded_customer_master, "得意先マスタ一覧"):
            reload_all_masters() # 保存成功したら再読み込み
            st.rerun() # ページをリフレッシュしてプレビューを更新
    else:
        st.error("ファイル名に「得意先マスタ一覧」が含まれていません。")
