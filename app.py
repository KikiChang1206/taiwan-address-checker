import streamlit as st
import pandas as pd
import io
import re

# 基本網頁設定
st.set_page_config(page_title="物流地址分類系統", layout="wide")
st.title("🚚 全台地址分類系統 (新竹/郵局)")

# --- 地址判斷邏輯優化 ---
def classify_address(address):
    if pd.isna(address):
        return "無鄉鎮"
    
    # 統一將「台」轉為「臺」，去除前後空白
    addr_str = str(address).replace("台", "臺").strip()
    
    # 1. 判斷【轉郵局】：離島區域、i郵箱、郵政信箱
    islands = ["澎湖", "金門", "連江", "馬祖", "蘭嶼", "綠島", "琉球"]
    post_keywords = ["i郵箱", "郵政信箱", "PO BOX", "郵局"]
    
    if any(island in addr_str for island in islands) or any(key in addr_str for key in post_keywords):
        return "轉郵局"
    
    # 2. 判斷【有鄉鎮】：優化正則表達式，偵測全台灣行政區特徵
    # 只要地址中包含 縣/市 加上 鄉/鎮/市/區，或是直接出現特定的區名
    pattern = r"([臺台].+[縣市].+[鄉鎮市區])|(.+[縣市][鄉鎮市區])|(.+[鄉鎮市區])"
    
    # 取消 [:10] 的限制，搜尋整個字串但以關鍵字為主
    if re.search(pattern, addr_str):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- Excel 導出格式設定 ---
def to_excel(df_to_save):
    output = io.BytesIO()
    final_df = df_to_save.copy()
    
    if 'category' in final_df.columns:
        final_df = final_df.drop(columns=['category'])
    
    # 補零邏輯 (針對所有包含「電話」字眼的欄位自動優化)
    for col in final_df.columns:
        if "電話" in col:
            final_df[col] = final_df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            final_df[col] = final_df[col].apply(lambda x: x.zfill(10) if (len(x) == 9 and x.startswith('9')) else x)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # --- 核心修改：防止文字溢出與重疊 ---
        style_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'num_format': '@',
            'text_wrap': False,   # 不自動換行
        })
        
        # 針對可能溢出的欄位，強制截斷或收縮
        overflow_prevent_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'fill',      # 關鍵：長文字會被截斷在格線內
            'valign': 'vcenter',
            'num_format': '@'
        })
        
        header_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': False,
            'align': 'left',
            'valign': 'vcenter',
            'border': 0
        })

        num_cols = len(final_df.columns)
        if num_cols > 0:
            # 預設套用「填充」格式，防止文字飄到隔壁
            worksheet.set_column(0, num_cols - 1, 12, overflow_prevent_format)
            for col_num, value in enumerate(final_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

    return output.getvalue()

# --- Streamlit UI 邏輯 ---
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None

uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(uploaded_file, engine=engine, dtype=str)
        
        # 彈性檢查地址欄位名稱
        addr_col = next((c for c in df.columns if "地址" in c), None)
        
        if not addr_col:
            st.error("❌ 錯誤：找不到包含『地址』字眼的欄位。")
        else:
            if not st.session_state['processed']:
                st.info(f"檔案已就緒，偵測到地址欄位：『{addr_col}』，共 {len(df)} 筆。")
            
            if st.button("🚀 執行分類並導出"):
                with st.spinner('處理中...'):
                    df['category'] = df[addr_col].apply(classify_address)
                    st.session_state['df_result'] = df
                    st.session_state['processed'] = True

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

                st.write("### 📥 下載分類結果")
                dl1, dl2, dl3 = st.columns(3)
                with dl1: st.download_button("📥 下載：轉郵局", to_excel(df_post), "轉郵局.xlsx")
                with dl2: st.download_button("📥 下載：轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮.xlsx")
                with dl3: st.download_button("📥 下載：轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮.xlsx")

    except Exception as e:
        st.error(f"系統異常：{e}")
