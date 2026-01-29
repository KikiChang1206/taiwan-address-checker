import streamlit as st
import pandas as pd
import io
import re

# 基本網頁設定
st.set_page_config(page_title="物流地址分類系統", layout="wide")
st.title("🚚 全台地址分類系統 (新竹/郵局)")

# --- 地址判斷邏輯 ---
def classify_address(address):
    if pd.isna(address):
        return "無鄉鎮"
    
    # 統一將「台」轉為「臺」，去除前後空白
    addr_str = str(address).replace("台", "臺").strip()
    
    # 1. 判斷【轉郵局】：離島區域、i郵箱、郵政信箱
    islands = ["澎湖", "金門", "連江", "馬祖", "蘭嶼", "綠島", "琉球"]
    post_keywords = ["i郵箱", "郵政信箱", "PO BOX"]
    
    if any(island in addr_str for island in islands) or any(key in addr_str for key in post_keywords):
        return "轉郵局"
    
    # 2. 判斷【縣市鄉鎮】：檢查前 10 個字是否包含行政區關鍵字
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[縣市])|(.+[鄉鎮市區])"
    if re.search(pattern, addr_str[:10]):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- Excel 導出格式設定 (Arial 10, 寬度 8.09, 透明無填滿) ---
def to_excel(df_to_save):
    output = io.BytesIO()
    # 移除分類用的輔助欄位
    final_df = df_to_save.drop(columns=['category']) if 'category' in df_to_save.columns else df_to_save
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 定義內容格式 (完全不指定 pattern 和 bg_color 即為透明無填滿)
        style_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # 定義標題格式
        header_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter'
        })

        # 取得欄位數量
        num_cols = len(final_df.columns)
        
        if num_cols > 0:
            # 設定所有欄位的寬度為 8.09，並套用透明格式
            worksheet.set_column(0, num_cols - 1, 8.09, style_format)
            
            # 重新寫入標題列以套用加粗格式
            for col_num, value in enumerate(final_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
        # 隱藏工作表背景格線
        worksheet.hide_gridlines(2)

    return output.getvalue()

# --- Streamlit UI 邏輯 ---

# 初始化狀態，避免下載時重整
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None

# 檔案上傳
uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # 根據副檔名選擇讀取引擎
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(uploaded_file, engine=engine)
        
        # 直接檢查標題
        if "收件人地址" not in df.columns:
            st.error("❌ 錯誤：找不到標題為『收件人地址』的欄位，請檢查檔案內容。")
        else:
            if not st.session_state['processed']:
                st.info(f"檔案已就緒，共 {len(df)} 筆。請點擊按鈕執行分類。")
            
            if st.button("🚀 執行分類並導出"):
                with st.spinner('分類處理中...'):
                    df['category'] = df["收件人地址"].apply(classify_address)
                    st.session_state['df_result'] = df
                    st.session_state['processed'] = True

            # 顯示處理結果
            if st.session_state['processed']:
                res_df = st.session_state['df_result']
                df_post = res_df[res_df['category'] == "轉郵局"]
                df_ok = res_df[res_df['category'] == "有鄉鎮"]
                df_no = res_df[res_df['category'] == "無鄉鎮"]

                st.write("---")
                st.subheader("📊 分類統計")
                c1, c2, c3 = st.columns(3)
                c1.metric("📮 轉郵局", f"{len(df_post)} 筆")
                c2.metric("🏠 轉新竹_有鄉鎮", f"{len(df_ok)} 筆")
                c3.metric("⚠️ 轉新竹_無鄉鎮", f"{len(df_no)} 筆")

                st.write("### 📥 下載分類結果 (Arial 10, 透明底)")
                dl1, dl2, dl3 = st.columns(3)
                
                with dl1:
                    st.download_button("📥 下載：轉郵局", to_excel(df_post), "轉郵局.xlsx", key="btn_p")
                with dl2:
                    st.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx", key="btn_ok")
                with dl3:
                    st.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx", key="btn_no")

    except Exception as e:
        st.error(f"系統發生異常：{e}")
else:
    st.session_state['processed'] = False
    st.session_state['df_result'] = None
    st.info("請上傳 Excel 檔案開始作業。")
