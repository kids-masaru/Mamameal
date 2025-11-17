import streamlit as st
import io
import os
from openpyxl import load_workbook
import pdfplumber

# st.set_page_config は削除 (streamlit_app.py で設定済みのため)

# --- シール/その他PDF処理関数 ---
def process_other_pdf_to_seal_template(pdf_bytes_io, existing_seal_path):
    """
    PDFの全テキストを抽出し、既存のseal.xlsxの最初のシートに貼り付ける
    """
    # 既存のseal.xlsxを読み込む
    wb = load_workbook(existing_seal_path)
    
    # 最初のシートを取得
    ws = wb.worksheets[0]
    
    # 既存のデータをクリア (1行目から最大行まで削除)
    if ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    current_row = 1 # Excelに書き込む現在の行番号
    
    # --- ▼▼ ここからロジックを大幅に変更 ▼▼ ---
    # extract_tables() の代わりに extract_text(layout=True) を使用
    
    with pdfplumber.open(pdf_bytes_io) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            
            # layout=True で、見た目のレイアウト（インデント等）を保持したままテキストを抽出
            page_text = page.extract_text(layout=True)
            
            if not page_text:
                continue

            # 抽出したテキストを1行ずつExcelに書き込む
            lines = page_text.split('\n')
            for line in lines:
                # 1列目 (A列) に行データを書き込む
                ws.cell(row=current_row, column=1, value=line)
                current_row += 1
                
            # ページ間に空行を1行入れる (見やすくするため)
            if page_number < len(pdf.pages):
                current_row += 1

    # --- ▲▲ ここまでロジックを大幅に変更 ▲▲ ---

    # データを書き込めなかった場合
    if current_row == 1:
        ws.cell(row=1, column=1, value="PDFからテキストを抽出できませんでした。")

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
        
        # アップロードされたPDFの元のファイル名（拡張子なし）を取得
        original_pdf_name = os.path.splitext(uploaded_pdf.name)[0]
        
        st.download_button(
            label=f"▼ seal_{original_pdf_name}.xlsx ダウンロード",
            data=modified_seal_bytes,
            file_name=f"seal_{original_pdf_name}.xlsx", # ファイル名を動的に変更
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"シール/その他PDF処理中にエラーが発生しました: {str(e)}")
        if show_debug:
            st.exception(e)
