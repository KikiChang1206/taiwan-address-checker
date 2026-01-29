import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="物流地址自動分流系統", layout="wide")
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
    
    # 2. 判斷是否有【縣市鄉鎮】
    # 檢查前 10 個字是否包含 縣/市/鄉/鎮/區
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[縣市])|(.+[鄉鎮市區])"
    if re.search(pattern, addr_str[:10]):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- UI 介面 ---
uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(uploaded_file, engine=engine)
        st.success(f"檔案讀取成功！共 {len(df)} 筆資料")
        
        all_cols = df.columns.tolist()
        default_index = all_cols.index("收件人地址") if "收件人地址" in all_cols else 0
        target_col = st.selectbox("請確認地址欄位：", all_cols, index=default_index)

        if st.button("執行分類"):
            # 執行分類
            df['category'] = df[target_col].apply(classify_address)
            
            df_post = df[df['category'] == "轉郵局"]
            df_ok = df[df['category'] == "有鄉鎮"]
            df_no = df[df['category'] == "無鄉鎮"]

            st.write("### 📊 分類統計")
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("轉郵局 (離島/i郵箱)", len(df_post))
            col_b.metric("轉新竹_有鄉鎮", len(df_ok))
            col_c.metric("轉新竹_無鄉鎮", len(df_no))

            # 下載 Function
            def to_excel(df_to_save):
                output = io.BytesIO()
                # 移除分類輔助欄位再儲存
                final_df = df_to_save.drop(columns=['category']) if 'category' in df_to_save.columns else df_to_save
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False)
                return output.getvalue()

            st.write("---")
            st.write("### 📥 下載分類結果")
            dl_col1, dl_col2, dl_col3 = st.columns(3)
            
            with dl_col1:
                st.download_button("📥 下載：轉郵局", to_excel(df_post), "轉郵局.xlsx")
            with dl_col2:
                st.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx")
            with dl_col3:
                st.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx")
                
    except Exception as e:
        st.error(f"錯誤：{e}")
