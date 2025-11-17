import streamlit as st
import io
import os
from openpyxl import load_workbook
import pdfplumber

# st.set_page_config は削除 (streamlit_app.py で設定済みのため)

# --- シール/その他PDF処理関数 ---
def process_other_pdf_to_seal_template(pdf_bytes_io, existing_seal_path):
    """
    seal.xlsxを読み込み、シートを2枚に分けてPDFデータを貼り付ける
    - 貼り付け1: 全文字の生データ (フォントサイズ確認用)
    - 貼り付け2: シンプルな全テキスト (基本情報)
    """
    # 既存のseal.xlsxを読み込む
    wb = load_workbook(existing_seal_path)
    
    # --- シートの準備 ---
    
    # 1. 1枚目のシートを「貼り付け1」にリネーム
    ws1 = wb.worksheets[0]
    ws1.title = "貼り付け1"
    
    # 2. 2枚目のシート「貼り付け2」を作成（または取得）
    if "貼り付け2" in wb.sheetnames:
        ws2 = wb["貼り付け2"]
    else:
        # 2番目の位置 (index=1) にシートを作成
        ws2 = wb.create_sheet(title="貼り付け2", index=1)

    # 3. 両方のシートの既存データをクリア
    if ws1.max_row > 0:
        ws1.delete_rows(1, ws1.max_row)
    if ws2.max_row > 0:
        ws2.delete_rows(1, ws2.max_row)

    # --- PDFからのデータ抽出 ---
    
    all_char_data = []  # 貼り付け1用 (文字の生データ)
    all_text_lines = [] # 貼り付け2用 (シンプルなテキスト)
    
    # 貼り付け1用のヘッダーを追加
    all_char_data.append(["Text", "FontSize", "x0", "y0"])

    with pdfplumber.open(pdf_bytes_io) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            
            # 【貼り付け1用】全文字の生データを抽出 (page.chars)
            # これで「大きな文字」がテキストとして存在するか確認
            for char in page.chars:
                all_char_data.append([
                    char.get("text"),
                    round(char.get("size"), 2) if char.get("size") else 0, # フォントサイズ
                    round(char.get("x0"), 2), # X座標
                    round(char.get("y0"), 2)  # Y座標
                ])
            
            # 【貼り付け2用】ページ全体のシンプルなテキストを抽出 (extract_text)
            page_text = page.extract_text()
            if page_text:
                all_text_lines.extend(page_text.split('\n'))

            # ページ間に区切りを入れる
            if page_number < len(pdf.pages):
                all_char_data.append(["-" * 10, f"Page {page_number+1}", "-" * 10, "-" * 10])
                all_text_lines.append(f"--- ページ {page_number+1} ---")


    # --- データをExcelに書き込み ---

    # 1. 「貼り付け1」にテーブルデータを書き込む
    if len(all_char_data) <= 1: # ヘッダーのみの場合
        ws1.cell(row=1, column=1, value="文字データを抽出できませんでした。")
    else:
        for r_idx, row_data in enumerate(all_char_data, start=1):
            if row_data:
                for c_idx, cell_data in enumerate(row_data, start=1):
                    cell_value = cell_data if cell_data is not None else ""
                    ws1.cell(row=r_idx, column=c_idx, value=cell_value)

    # 2. 「貼り付け2」に全テキストデータを書き込む
    if not all_text_lines:
        ws2.cell(row=1, column=1, value="PDFからテキストを抽出できませんでした。")
    else:
        for r_idx, line in enumerate(all_text_lines, start=1):
            ws2.cell(row=r_idx, column=1, value=line)

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
