import streamlit as st
import pandas as pd
import io
import re

# 基本網頁設定
st.set_page_config(page_title="物流地址分類系統 v2", layout="wide")
st.title("🚚 全台地址分類系統 (格式優化版)")

# --- 地址判斷邏輯：解決「新莊區」等無縣市開頭問題 ---
def classify_address(address):
    if pd.isna(address):
        return "無鄉鎮"
    
    addr_str = str(address).replace("台", "臺").strip()
    
    # 1. 判斷【轉郵局】：離島區域、i郵箱、郵政信箱
    islands = ["澎湖", "金門", "連江", "馬祖", "蘭嶼", "綠島", "琉球"]
    post_keywords = ["i郵箱", "郵政信箱", "PO BOX", "郵局"]
    
    if any(island in addr_str for island in islands) or any(key in addr_str for key in post_keywords):
        return "轉郵局"
    
    # 2. 判斷【有鄉鎮】：不限字數位置，偵測台灣常見行政區特徵
    # 包含：XX區、XX鄉、XX鎮、XX市
    pattern = r"(.+[縣市].+[鄉鎮市區])|(.+[鄉鎮市區])"
    
    if re.search(pattern, addr_str):
        return "有鄉鎮"
    
    return "無鄉鎮"

# --- Excel 導出格式設定：解決文字溢出蓋到隔壁欄位 ---
def to_excel(df_to_save):
    output = io.BytesIO()
    final_df = df_to_save.copy()
    
    if 'category' in final_df.columns:
        final_df = final_df.drop(columns=['category'])
    
    # 處理電話補零
    for col in final_df.columns:
        if "電話" in col or "連" in col: # 增加對「連」絡電話的偵測
            final_df[col] = final_df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            final_df[col] = final_df[col].apply(lambda x: x.zfill(10) if (len(x) == 9 and x.startswith('9')) else x)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # --- 核心修改：解決溢出問題 ---
        # shrink: True 會讓長文字自動縮小在格線內，不影響閱讀也不會蓋到旁邊
        style_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'num_format': '@',
            'shrink': True,  # 關鍵：自動收縮文字，防止溢出
            'border': 0
        })
        
        # 標題格式
        header_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': False,
            'align': 'left',
            'valign': 'vcenter',
            'border': 0,
            'num_format': '@'
        })

        num_cols = len(final_df.columns)
        if num_cols > 0:
            # 設定標準欄寬 10 (比 8.09 稍微寬一點更美觀)
            worksheet.set_column(0, num_cols - 1, 10, style_format)
            
            # 強制標題列不使用收縮
            for col_num, value in enumerate(final_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

    return output.getvalue()

# --- Streamlit UI 邏輯 ---
if 'processed' not in st.session_state:
    st.session_state['processed'] = False
if 'df_result' not in st.session_state:
    st.session_state['df_result'] = None

uploaded_file = st.file_uploader("請上傳您的原始 Excel 檔案", type=["xls", "xlsx"])

if uploaded_file:
    try:
        engine = 'xlrd' if uploaded_file.name.endswith('.xls') else 'openpyxl'
        # 讀取時強制轉字串，守住電話號碼的 0
        df = pd.read_excel(uploaded_file, engine=engine, dtype=str)
        
        # 自動搜尋地址欄位
        addr_col = next((c for c in df.columns if "地址" in c or "地" in c), None)
        
        if not addr_col:
            st.error("❌ 找不到地址欄位，請確認標題是否有『地址』或『地』字眼。")
        else:
            if not st.session_state['processed']:
                st.info(f"成功讀取檔案！共 {len(df)} 筆資料。")
            
            if st.button("🚀 開始分類並修復格式"):
                with st.spinner('優化中...'):
                    df['category'] = df[addr_col].apply(classify_address)
                    st.session_state['df_result'] = df
                    st.session_state['processed'] = True

            if st.session_state['processed']:
                res_df = st.session_state['df_result']
                df_post = res_df[res_df['category'] == "轉郵局"]
                df_ok = res_df[res_df['category'] == "有鄉鎮"]
                df_no = res_df[res_df['category'] == "無鄉鎮"]

                st.divider()
                st.subheader("📊 分類結果摘要")
                c1, c2, c3 = st.columns(3)
                c1.metric("📮 轉郵局", f"{len(df_post)} 筆")
                c2.metric("🏠 轉新竹_有鄉鎮", f"{len(df_ok)} 筆")
                c3.metric("⚠️ 轉新竹_無鄉鎮", f"{len(df_no)} 筆")

                st.write("### 📥 下載優化後的檔案")
                dl1, dl2, dl3 = st.columns(3)
                with dl1: st.download_button("📥 轉郵局", to_excel(df_post), "轉郵局_已修復.xlsx")
                with dl2: st.download_button("📥 轉新竹_有鄉鎮", to_excel(df_ok), "轉新竹_有鄉鎮_已修復.xlsx")
                with dl3: st.download_button("📥 轉新竹_無鄉鎮", to_excel(df_no), "轉新竹_無鄉鎮_已修復.xlsx")

    except Exception as e:
        st.error(f"系統發生錯誤：{e}")
