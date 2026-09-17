import streamlit as st
import pandas as pd

st.set_page_config(page_title="卡美問題查詢", layout="centered")

# --- 讀取與快取資料 ---
@st.cache_data
def load_data():
    df = pd.read_excel("RCA.xlsx", dtype=str)
    df.fillna("無資料", inplace=True)
    return df

try:
    df = load_data()
except FileNotFoundError:
    st.error("找不到 RCA.xlsx 檔案！請確認它是否放在跟 app.py 同一個資料夾內。")
    st.stop()

st.header("🔍 故障排除查詢")

# ==========================================
# 📱 核心過濾區
# ==========================================
with st.container(border=True):
    unique_stations = [x for x in df["STATION"].unique() if x != "無資料"]
    station_options = ["ALL"] + unique_stations
    
    selected_station = st.selectbox("📌 選擇站別", station_options)
    
    if selected_station == "ALL":
        base_df = df
    else:
        base_df = df[df["STATION"] == selected_station]

    search_method = st.radio("🔍 第一步：選擇查詢方式", ["使用 BIN_CODE", "使用 BIN"], horizontal=True)

    filtered_df = pd.DataFrame()
    selected_sub_bin = None
    
    display_bin_code = ""
    display_bin = ""

    if search_method == "使用 BIN_CODE":
        unique_bin_codes = [x for x in base_df["BIN_CODE"].unique() if x != "無資料"]
        selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN_CODE", unique_bin_codes)
        
        if selected_val:
            filtered_df = base_df[base_df["BIN_CODE"] == selected_val]
            associated_bins = [x for x in filtered_df["BIN"].unique() if x != "無資料"]
            
            display_bin_code = selected_val
            display_bin = ', '.join(associated_bins) if associated_bins else "(無對應紀錄)"
            
            st.caption(f"💡 對應 BIN: {display_bin}")
            
    else:
        unique_bins = [x for x in base_df["BIN"].unique() if x != "無資料"]
        selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN", unique_bins)
        
        if selected_val:
            filtered_df = base_df[base_df["BIN"] == selected_val]
            associated_bin_codes = [x for x in filtered_df["BIN_CODE"].unique() if x != "無資料"]
            
            display_bin = selected_val
            display_bin_code = ', '.join(associated_bin_codes) if associated_bin_codes else "(無對應紀錄)"
            
            st.caption(f"💡 對應 BIN_CODE: {display_bin_code}")
            st.caption(f"💡 BIN 全文: {display_bin}")

    if not filtered_df.empty:
        unique_sub_bins = filtered_df["SUB_BIN"].unique()
        selected_sub_bin = st.selectbox("📑 第三步：請選擇 SUB_BIN", unique_sub_bins)

# ==========================================
# 🖥️ 顯示結果區域
# ==========================================
if not filtered_df.empty and selected_sub_bin:
    st.divider() 
    
    st.markdown(f"**🏷️ BIN_CODE:** {display_bin_code}")
    st.markdown(f"**🏷️ BIN 全文:**  \n{display_bin}")
    
    if selected_sub_bin != "無資料":
        st.markdown(f"**🏷️ SUB_BIN 全文:**  \n{selected_sub_bin}")
    
    final_df = filtered_df[filtered_df["SUB_BIN"] == selected_sub_bin]
    st.markdown(f"### 💡 找到 {len(final_df)} 筆解決方案")
    
    for index, row in final_df.iterrows():
        with st.container(border=True):
            
            # ★ 修改處：先用 replace('\\n', '\n') 處理純文字的字元，再替換成 Markdown 的換行格式
            cause_text = str(row['Possible Cause']).replace('\\n', '\n').replace('\n', '  \n')
            solution_text = str(row['Solution']).replace('\\n', '\n').replace('\n', '  \n')
            
            st.error(f"**🚨 可能原因 (Cause):**  \n{cause_text}")
            st.success(f"**✅ 解決方案 (Solution):**  \n{solution_text}")
            
            meta_info = []
            if row['Ref Log'] != "無資料":
                meta_info.append(f"**Log:** {row['Ref Log']}")
            if row['REV'] != "無資料":
                meta_info.append(f"**REV:** {row['REV']}")
            
            if meta_info:
                st.caption(" | ".join(meta_info))
