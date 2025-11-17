import streamlit as st
import io
import os
from openpyxl import load_workbook
import pdfplumber

# st.set_page_config は削除 (streamlit_app.py で設定済みのため)

# --- シール/その他PDF処理関数 ---
def process_other_pdf_to_seal_template(pdf_bytes_io, existing_seal_path):
    """
    数出表以外のPDFを処理し、既存のseal.xlsxの最初のシートに全テーブルデータを貼り付ける
    """
    # 既存のseal.xlsxを読み込む
    wb = load_workbook(existing_seal_path)
    
    # 最初のシートを取得
    ws = wb.worksheets[0]
    
    # 既存のデータをクリア (1行目から最大行まで削除)
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    all_rows_data = []
    
    # pdfplumberでPDFを開き、全ページの全テーブルを抽出
    with pdfplumber.open(pdf_bytes_io) as pdf:
        for page in pdf.pages:
            # --- ▼▼ 修正点 ▼▼ ---
            # "strategy": "text" を正しい場所（table_settingsの外）に指定
            tables = page.extract_tables(table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text"
            })
            
            # もし上記でテーブルが見つからなかった場合、"text" 戦略を試す
            if not tables:
                 tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3, # 少しのズレを許容
                })

            # --- ▲▲ 修正点 ▲▲ ---
            
            if tables:
                for table in tables:
                    # tableはネストされたリスト (list of lists)
                    all_rows_data.extend(table)
    
    # 抽出したデータをシートに書き込む
    if not all_rows_data:
        # 最終手段として、テーブルが見つからない場合はページ全体のテキストを抽出
        ws.cell(row=1, column=1, value="テーブル抽出失敗。ページ全体のテキストを抽出します。")
        current_row = 2
        with pdfplumber.open(pdf_bytes_io) as pdf_text:
             for page in pdf_text.pages:
                 page_text = page.extract_text()
                 if page_text:
                     for line in page_text.split('\n'):
                         ws.cell(row=current_row, column=1, value=line)
                         current_row += 1

    else:
        for r_idx, row_data in enumerate(all_rows_data, start=1):
            if row_data: # 行データが空でないことを確認
                for c_idx, cell_data in enumerate(row_data, start=1):
                    # セルデータがNoneの場合は空文字に変換
                    cell_value = cell_data if cell_data is not None else ""
                    ws.cell(row=r_idx, column=c_idx, value=cell_value)

    # 変更をバイトデータとして保存
    output_excel = io.BytesIO()
    wb.save(output_excel)
    return output_excel.getvalue()

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


st.markdown('<p class="custom-title">シール貼付 PDF変換ツール</p>', unsafe_allow_html=True)

uploaded_pdf = st.file_uploader("処理するシールPDFファイルをアップロードしてください", type="pdf", label_visibility="collapsed")

if uploaded_pdf is not None:
    
    pdf_bytes_io = io.BytesIO(uploaded_pdf.getvalue())
    
    st.info("📜 シール/その他PDFとして処理します...")
    seal_template_path = "seal.xlsx" # フォルダにある既存のファイル
    
    if not os.path.exists(seal_template_path):
        st.error(f"必要なテンプレートファイル '{seal_template_path}' が見つかりません。")
        st.stop()
    
    try:
        with st.spinner(f"'{seal_template_path}' にPDFデータを書き込み中..."):
            # 新しい関数を呼び出し、変更されたExcelのバイトデータを取得
            modified_seal_bytes = process_other_pdf_to_seal_template(pdf_bytes_io, seal_template_path)
        
        st.success(f"✅ 処理が完了しました！")
        
        # --- ▼▼ 修正点 (ファイル名) ▼▼ ---
        # アップロードされたPDFの元のファイル名（拡張子なし）を取得
        original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
        
        st.download_button(
            label=f"▼ seal_{original_pdf_name}.xlsx ダウンロード",
            data=modified_seal_bytes,
            file_name=f"seal_{original_pdf_name}.xlsx", # ファイル名を動的に変更
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # --- ▲▲ 修正点 (ファイル名) ▲▲ ---
        
    except Exception as e:
        st.error(f"シール/その他PDF処理中にエラーが発生しました: {str(e)}")
        if show_debug:
            st.exception(e)
