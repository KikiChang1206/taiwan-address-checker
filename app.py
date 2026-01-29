import streamlit as st
import pandas as pd
import io
import re

st.title("📦 物流地址自動分類系統")

# 定義判斷邏輯
def is_valid_address(address):
    if pd.isna(address): return False
    addr = str(address).replace("台", "臺").strip()
    # 檢查前10個字內是否出現縣、市、鄉、鎮、區等關鍵字
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[縣市])|(.+[鄉鎮市區])"
    return bool(re.search(pattern, addr[:10]))

uploaded_file = st.file_uploader("上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    # 支援新舊版 Excel
    engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
    df = pd.read_excel(uploaded_file, engine=engine)
    
    # 自動尋找 Z 欄位（第 26 欄）
    cols = df.columns.tolist()
    z_col = cols[25] if len(cols) >= 26 else cols[-1]
    target_col = st.selectbox("請確認地址欄位：", cols, index=cols.index(z_col))

    if st.button("開始分類"):
        mask = df[target_col].apply(is_valid_address)
        df_ok = df[mask]
        df_no = df[~mask]

        st.success(f"分類完成！有縣市：{len(df_ok)} 筆 / 無縣市：{len(df_no)} 筆")

        def to_excel(df_data):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                df_data.to_excel(writer, index=False)
            return out.getvalue()

        c1, c2 = st.columns(2)
        c1.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx")
        c2.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx")
