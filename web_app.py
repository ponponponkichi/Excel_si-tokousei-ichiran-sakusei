import streamlit as st
import pandas as pd
import datetime
import io
import openpyxl
import xlrd

# --------------------------------------------------------------------
# ■ ページ基本設定
# --------------------------------------------------------------------
st.set_page_config(
    page_title="複数Excelファイルのシート構成一覧作成 Web版",
    layout="centered"
)

# 大タイトル（上下の余白を極限までカット）
st.markdown("""
    <div style="margin-top: 5px; margin-bottom: 0px;">
        <h1 style="font-size: 2.0rem; color: var(--text-color); font-weight: 700; border: none; padding: 0; margin: 0;">
            複数Excelファイルのシート構成一覧作成ツール
        </h1>
    </div>
""", unsafe_allow_html=True)
st.caption("複数Excelファイルのシート名一覧を作成し、非表示シートを可視化します。")


# --------------------------------------------------------------------
# ■ 1. ファイル選択セクション
# --------------------------------------------------------------------
st.markdown("""
    <div style="border-left: 5px solid #107C41; padding-left: 10px; margin-top: 25px; margin-bottom: 10px;">
        <h2 style="font-size: 1.3rem; color: var(--text-color); font-weight: 700; margin: 0; padding: 0; border: none;">
            1. 対象ファイルのアップロード
        </h2>
    </div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "対象のExcel（.xlsx, .xls）ファイルをまとめて選択、またはドロップしてください",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# --------------------------------------------------------------------
# ■ 2. 処理実行 & ダウンロード
# --------------------------------------------------------------------
st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)

if st.button("処理を実行する", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("ファイルがアップロードされていません。")
    else:
        try:
            with st.spinner("処理を実行中..."):
                all_files_data = []
                hidden_cells_info = []
                processed_files_count = 0

                # 各Excelファイルを処理
                for file in uploaded_files:
                    file_name = file.name
                    sheet_info = []

                    try:
                        # .xlsx ファイルの処理
                        if file_name.lower().endswith('.xlsx'):
                            workbook = openpyxl.load_workbook(file, read_only=True)
                            for sheet in workbook.worksheets:
                                is_hidden = sheet.sheet_state != 'visible'
                                sheet_info.append((sheet.title, is_hidden))
                        
                        # .xls ファイルの処理
                        elif file_name.lower().endswith('.xls'):
                            file_contents = file.read()
                            workbook = xlrd.open_workbook(file_contents=file_contents, on_demand=True)
                            for i in range(len(workbook.sheet_names())):
                                sheet = workbook.sheet_by_index(i)
                                is_hidden = sheet.visibility != 0
                                sheet_info.append((sheet.name, is_hidden))
                        else:
                            continue

                        # 1行分のデータを構築
                        row_data = {'ファイル名': file_name}
                        
                        for i, (sheet_name, is_hidden) in enumerate(sheet_info):
                            col_name = f'シート{i+1}'
                            row_data[col_name] = sheet_name
                            
                            if is_hidden:
                                hidden_cells_info.append({
                                    'row_idx': processed_files_count, 
                                    'col_name': col_name,
                                    'value': sheet_name
                                })
                        
                        all_files_data.append(row_data)
                        processed_files_count += 1

                    except Exception as e:
                        st.error(f"エラー: {file_name} の処理中にエラーが発生しました: {e}")
                
                if not all_files_data:
                    raise ValueError("有効なExcelファイルを1つも処理できませんでした。")

                # --- データフレーム作成と保存 ---
                result_df = pd.DataFrame(all_files_data)

                # 列の並び順を整理 (ファイル名, シート1, シート2...)
                sheet_cols = [col for col in result_df.columns if col.startswith('シート')]
                sheet_cols.sort(key=lambda x: int(x.replace('シート', '')))
                
                final_columns = ['ファイル名'] + sheet_cols
                result_df = result_df.reindex(columns=final_columns)

                # メモリ上への書き出し (Excel形式)
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    sheet_name = 'シート構成一覧'
                    result_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    # 書式設定: グレー背景、濃いグレー文字
                    hidden_format = workbook.add_format({
                        'bg_color': '#D3D3D3',
                        'font_color': '#5A5A5A'
                    })

                    col_map = {name: i for i, name in enumerate(result_df.columns)}
                    
                    # 非表示セルの書式適用
                    for info in hidden_cells_info:
                        row_idx = info['row_idx']
                        col_name = info['col_name']
                        value = info['value']
                        
                        if col_name in col_map:
                            col_idx = col_map[col_name]
                            excel_row = row_idx + 1
                            worksheet.write(excel_row, col_idx, value, hidden_format)

                    # カラム幅の自動調整 (簡易実装)
                    for i, col in enumerate(result_df.columns):
                        max_len = len(str(col))
                        column_data = result_df[col].astype(str)
                        if not column_data.empty:
                            max_data_len = column_data.map(len).max()
                            if max_data_len > max_len:
                                max_len = max_data_len
                        worksheet.set_column(i, i, max_len + 2)

                now = datetime.datetime.now().strftime("%Y%m%d")
                filename = f"シート構成一覧_{now}.xlsx"
                download_data = excel_buffer.getvalue()

                st.success(f"処理完了！ {processed_files_count}個のファイルのシート構成を一覧化しました。")
                
                st.download_button(
                    label="一覧データをダウンロードする",
                    data=download_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"処理中にエラーが発生しました:\n{e}")

# --------------------------------------------------------------------
# ■ 説明インフォメーション
# --------------------------------------------------------------------
st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
with st.expander("詳細な仕様・前提条件を確認する"):
    st.markdown("""
    **【出力結果のイメージ】**
    * A列: ファイル名
    * B列: 1番目のシート名  
    * C列: 2番目のシート名  
    * 以降、D列、E列とシート順に続きます。
    
    **【仕様】**
    * デスクトップ版の「フォルダ指定」から「複数ファイル選択」に変更されています。
    * 非表示シートは、抽出されたExcelファイル上で「グレーの背景色」で表示されます。
    * パスワード保護されているExcelファイルは読み込めません。
    """)