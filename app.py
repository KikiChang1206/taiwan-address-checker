import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="物流地址自動分類系統", layout="wide")
st.title("🚚 全台地址分類系統 (新竹/郵局)")

# --- 核心檢查邏輯 ---
def classify_address(address):
    if pd.isna(address):
        return "無鄉鎮"
    
    addr_str = str(address).replace("台", "臺").strip()
    
    # 1. 判斷是否為【轉郵局】 (離島、i郵箱、郵政信箱)
    islands = ["澎湖", "金門", "連江", "馬祖", "蘭嶼", "綠島", "琉球"]
    post_keywords = ["i郵箱", "郵政信箱", "PO BOX"]
    
    if any(island in addr_str for island in islands) or any(key in addr_str for key in post_keywords):
        return "轉郵局"
    
    # 2. 判斷是否有【縣市鄉鎮】 (檢查前 10 個字)
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[縣市])|(.+[鄉鎮市區])"
    if re.search(pattern, addr_str[:10]):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- 檔案下載 Function ---
def to_excel(df_to_save):
    output = io.BytesIO()
    # 移除分類輔助欄位再儲存
    final_df = df_to_save.drop(columns=['category']) if 'category' in df_to_save.columns else df_to_save
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
    return output.getvalue()

# --- UI 介面 ---
uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # 讀取檔案
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(uploaded_file, engine=engine)
        
        # 檢查是否有「收件人地址」欄位
        if "收件人地址" not in df.columns:
            st.error("❌ 錯誤：檔案中找不到『收件人地址』欄位，請檢查標題是否正確。")
        else:
            # 使用 Session State 來儲存分類結果，避免下載後消失
            if st.button("🚀 開始分類資料"):
                with st.spinner('處理中...'):
                    df['category'] = df["收件人地址"].apply(classify_address)
                    st.session_state['df_result'] = df
                    st.session_state['processed'] = True

            # 如果已經處理過，就顯示結果與下載按鈕
            if st.session_state.get('processed'):
                res_df = st.session_state['df_result']
                df_post = res_df[res_df['category'] == "轉郵局"]
                df_ok = res_df[res_df['category'] == "有鄉鎮"]
                df_no = res_df[res_df['category'] == "無鄉鎮"]

                st.write("---")
                st.write("### 📊 分類統計 (處理完成)")
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("轉郵局 (離島/i郵箱)", len(df_post))
                col_b.metric("轉新竹_有鄉鎮", len(df_ok))
                col_c.metric("轉新竹_無鄉鎮", len(df_no))

                st.write("### 📥 下載分類結果")
                dl_col1, dl_col2, dl_col3 = st.columns(3)
                
                with dl_col1:
                    st.download_button("📥 下載：轉郵局", to_excel(df_post), "轉郵局.xlsx", key="btn_post")
                with dl_col2:
                    st.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx", key="btn_ok")
                with dl_col3:
                    st.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx", key="btn_no")
                
                # 預覽部分資料
                st.write("---")
                st.write("🔍 分類結果預覽 (前 5 筆)：")
                st.dataframe(res_df[["收件人地址", "category"]].head())

    except Exception as e:
        st.error(f"系統發生異常：{e}")
else:
    # 當沒有檔案上傳時，重設狀態
    st.session_state['processed'] = False
    st.info("請上傳 Excel 檔案開始作業")
