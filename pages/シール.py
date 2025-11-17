import streamlit as st
import io
import os
from openpyxl import load_workbook
import pdfplumber

# --- ▼▼ 新しくインポート ▼▼ ---
import pytesseract
from pdf2image import convert_from_bytes
# --- ▲▲ 新しくインポート ▲▲ ---


# --- シール/その他PDF処理関数 ---
def process_other_pdf_to_seal_template(pdf_bytes_io, existing_seal_path):
    """
    seal.xlsxを読み込み、シートを2枚に分けてPDFデータを貼り付ける
    - 貼り付け1: OCRによる画像認識テキスト (★大きな文字用)
    - 貼り付け2: 従来のテキストデータ (★小さい文字用)
    """
    # 既存のseal.xlsxを読み込む
    wb = load_workbook(existing_seal_path)
    
    # --- シートの準備 ---
    ws1 = wb.worksheets[0]
    ws1.title = "貼り付け1" # OCR結果
    
    if "貼り付け2" in wb.sheetnames:
        ws2 = wb["貼り付け2"] # 従来のテキスト抽出
    else:
        ws2 = wb.create_sheet(title="貼り付け2", index=1)

    if ws1.max_row > 0: ws1.delete_rows(1, ws1.max_row)
    if ws2.max_row > 0: ws2.delete_rows(1, ws2.max_row)

    
    # --- ▼▼ 処理1: 【貼り付け1】OCRによる抽出 ▼▼ ---
    # (大きな文字が画像化されている場合に対応)
    ws1_current_row = 1
    try:
        # 1. PDF(バイトデータ)を画像のリストに変換
        images = convert_from_bytes(pdf_bytes_io.getvalue())
        
        if not images:
            ws1.cell(row=1, column=1, value="PDFを画像に変換できませんでした。")
        else:
            ws1.cell(row=1, column=1, value="--- OCR抽出開始 ---")
            ws1_current_row += 1
            
            for i, page_image in enumerate(images, 1):
                # 2. 画像から日本語テキストを抽出 (lang='jpn')
                ocr_text = pytesseract.image_to_string(page_image, lang='jpn', config='--psm 6')
                
                ws1.cell(row=ws1_current_row, column=1, value=f"--- ページ {i} (OCR) ---")
                ws1_current_row += 1
                
                if ocr_text:
                    # --- ▼▼ ここから修正 (貼り付け1) ▼▼ ---
                    # テキストを改行（\n）でリストに分割し、空白行を除外
                    lines_raw = ocr_text.split('\n')
                    lines = [line for line in lines_raw if line.strip()]
                    
                    # 4要素（4列）を1行として書き込む
                    num_lines = len(lines)
                    for i in range(0, num_lines, 4):
                        c1 = lines[i] if i < num_lines else None
                        c2 = lines[i+1] if (i+1) < num_lines else None
                        c3 = lines[i+2] if (i+2) < num_lines else None
                        c4 = lines[i+3] if (i+3) < num_lines else None
                        
                        ws1.cell(row=ws1_current_row, column=1, value=c1)
                        ws1.cell(row=ws1_current_row, column=2, value=c2)
                        ws1.cell(row=ws1_current_row, column=3, value=c3)
                        ws1.cell(row=ws1_current_row, column=4, value=c4)
                        ws1_current_row += 1 # 1行（4列）書くごとに行番号を増やす
                    # --- ▲▲ ここまで修正 (貼り付け1) ▲▲ ---
                else:
                    ws1.cell(row=ws1_current_row, column=1, value="(このページではOCRで文字を認識できませんでした)")
                    ws1_current_row += 1

    except Exception as e:
        ws1.cell(row=1, column=1, value=f"OCR処理中にエラーが発生しました: {str(e)}")
    # --- ▲▲ 処理1: OCR 完了 ▲▲ ---


    # --- ▼▼ 処理2: 【貼り付け2】従来のテキスト抽出 ▼▼ ---
    ws2_current_row = 1
    try:
        with pdfplumber.open(pdf_bytes_io) as pdf:
            if not pdf.pages:
                ws2.cell(row=1, column=1, value="PDFにページがありません。")
            else:
                for page_number, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text() # レイアウトなしのシンプルな抽出
                    
                    ws2.cell(row=ws2_current_row, column=1, value=f"--- ページ {page_number} (テキスト) ---")
                    ws2_current_row += 1
                    
                    if page_text:
                        # --- ▼▼ ここから修正 (貼り付け2) ▼▼ ---
                        # テキストを改行（\n）でリストに分割し、空白行を除外
                        lines_raw = page_text.split('\n')
                        lines = [line for line in lines_raw if line.strip()]
                        
                        # 4要素（4列）を1行として書き込む
                        num_lines = len(lines)
                        for i in range(0, num_lines, 4):
                            # 4つのデータを取得（足りない場合はNoneになる）
                            c1 = lines[i] if i < num_lines else None
                            c2 = lines[i+1] if (i+1) < num_lines else None
                            c3 = lines[i+2] if (i+2) < num_lines else None
                            c4 = lines[i+3] if (i+3) < num_lines else None
                            
                            ws2.cell(row=ws2_current_row, column=1, value=c1)
                            ws2.cell(row=ws2_current_row, column=2, value=c2)
                            ws2.cell(row=ws2_current_row, column=3, value=c3)
                            ws2.cell(row=ws2_current_row, column=4, value=c4)
                            ws2_current_row += 1 # 1行（4列）書くごとに行番号を増やす
                        # --- ▲▲ ここまで修正 (貼り付け2) ▲▲ ---
                    else:
                        ws2.cell(row=ws2_current_row, column=1, value="(このページではテキストを抽出できませんでした)")
                        ws2_current_row += 1
    
    except Exception as e:
        ws2.cell(row=1, column=1, value=f"テキスト抽出中にエラーが発生しました: {str(e)}")
        
    # --- ▲▲ 処理2: テキスト抽出 完了 ▲▲ ---

    # 変更をバイトデータとして保存
    output_excel = io.BytesIO()
    wb.save(output_excel)
    return output_excel.getvalue()

# --- ページ表示 (変更なし) ---

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
        with st.spinner(f"OCR処理とテキスト抽出を実行中... (時間がかかる場合があります)"):
            # 新しい関数を呼び出し、変更されたExcelのバイトデータを取得
            modified_seal_bytes = process_other_pdf_to_seal_template(pdf_bytes_io, seal_template_path)
        
        st.success(f"✅ 処理が完了しました！")
        
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
