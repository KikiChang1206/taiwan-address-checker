import streamlit as st
import pandas as pd
import io
import re

# 設定網頁標題與排版
st.set_page_config(page_title="物流地址自動分類系統", layout="wide")
st.title("🚚 全台地址分類系統 (新竹/郵局)")

# --- 核心檢查邏輯 ---
def classify_address(address):
    if pd.isna(address):
        return "無鄉鎮"
    
    # 統一將「台」轉為「臺」，並去除前後空白
    addr_str = str(address).replace("台", "臺").strip()
    
    # 1. 判斷是否為【轉郵局】 (離島、i郵箱、郵政信箱)
    islands = ["澎湖", "金門", "連江", "馬祖", "蘭嶼", "綠島", "琉球"]
    post_keywords = ["i郵箱", "郵政信箱", "PO BOX"]
    
    if any(island in addr_str for island in islands) or any(key in addr_str for key in post_keywords):
        return "轉郵局"
    
    # 2. 判斷是否有【縣市鄉鎮】 (利用正規表達式檢查地址前 10 個字)
    # 邏輯：檢查是否包含 縣市 或 鄉鎮市區 關鍵字
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[縣市])|(.+[鄉鎮市區])"
    if re.search(pattern, addr_str[:10]):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- 檔案下載 Function (含 Arial 10, 無框線設定) ---
def to_excel(df_to_save):
    output = io.BytesIO()
    # 移除分類輔助欄位
    final_df = df_to_save.drop(columns=['category']) if 'category' in df_to_save.columns else df_to_save
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 設定格式：Arial, 10號字, 無框線, 靠左對齊
        cell_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # 設定標題格式 (Arial 10, 加粗, 無框線)
        header_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'border': 0,
            'align': 'left',
            'valign': 'vcenter'
        })

        # 套用格式與設定欄位寬度
        for col_num, value in enumerate(final_df.columns.values):
            # 寫入標題列
            worksheet.write(0, col_num, value, header_format)
            # 套用內容格式到整欄，並設定預設寬度 25
            worksheet.set_column(col_num, col_num, 25, cell_format)
            
        # 隱藏 Excel 預設格線
        worksheet.hide_gridlines(2)

    return output.getvalue()

# --- UI 介面 ---
# 初始化 Session State (確保下載後結果不消失)
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None

uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # 判斷引擎：舊版 xls 使用 xlrd，新版 xlsx 使用 openpyxl
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(uploaded_file, engine=engine)
        
        # 檢查關鍵欄位
        if "收件人地址" not in df.columns:
            st.error("❌ 錯誤：檔案中找不到標題為『收件人地址』的欄位。")
        else:
            if not st.session_state['processed']:
                st.success(f"檔案讀取成功！共 {len(df)} 筆資料。請點擊下方按鈕開始分類。")
            
            if st.button("🚀 執行分類與格式化"):
                with st.spinner('正在分析地址並套用 Arial 格式...'):
                    df['category'] = df["收件人地址"].apply(classify_address)
                    st.session_state['df_result'] = df
                    st.session_state['processed'] = True

            # 顯示結果區域
            if st.session_state['processed']:
                res_df = st.session_state['df_result']
                df_post = res_df[res_df['category'] == "轉郵局"]
                df_ok = res_df[res_df['category'] == "有鄉鎮"]
                df_no = res_df[res_df['category'] == "無鄉鎮"]

                st.write("---")
                st.subheader("📊 分類統計結果")
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("📮 轉郵局 (離島/i郵箱)", f"{len(df_post)} 筆")
                col_b.metric("🏠 轉新竹_有鄉鎮", f"{len(df_ok)} 筆")
                col_c.metric("⚠️ 轉新竹_無鄉鎮", f"{len(df_no)} 筆")

                st.write("### 📥 下載分類檔案 (Arial 10 級字)")
                dl_col1, dl_col2, dl_col3 = st.columns(3)
                
                with dl_col1:
                    st.download_button("📥 下載：轉郵局", to_excel(df_post), "轉郵局.xlsx", key="btn_post")
                with dl_col2:
                    st.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx", key="btn_ok")
                with dl_col3:
                    st.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx", key="btn_no")
                
                # 資料預覽
                st.write("---")
                st.write("🔍 前 5 筆資料預覽：")
                st.dataframe(res_df[["收件人地址", "category"]].head())

    except Exception as e:
        st.error(f"系統發生異常，請確認檔案格式是否正確：{e}")
else:
    # 清空狀態以便上傳新檔案
    st.session_state['processed'] = False
    st.session_state['df_result'] = None
    st.info("💡 請上傳包含『收件人地址』標題的 Excel 檔案。")
