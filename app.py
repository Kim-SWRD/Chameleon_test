import streamlit as st
import pandas as pd

# 1. 將 layout 改為 "wide" (寬螢幕模式)，搭配側邊欄在電腦上視覺效果最好
st.set_page_config(page_title="卡美問題查詢", layout="wide", initial_sidebar_state="expanded")

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

# ==========================================
# 📱 側邊欄區域 (手機會自動收合，電腦會固定在左側)
# ==========================================
with st.sidebar:
    st.title("🛠️ 篩選條件")
    st.markdown("請依序選擇發生問題的站別與代碼。")
    
    # --- 步驟零：STATION 篩選 ---
    unique_stations = [x for x in df["STATION"].unique() if x != "無資料"]
    station_options = ["ALL"] + unique_stations
    
    selected_station = st.selectbox("📌 選擇站別 (STATION)", station_options)

    if selected_station == "ALL":
        base_df = df
    else:
        base_df = df[df["STATION"] == selected_station]

    st.divider()

    # --- 步驟一：選擇查詢方式 ---
    search_method = st.radio("🔍 第一步：選擇查詢方式", ["使用 BIN_CODE", "使用 BIN"])
    
    filtered_df = pd.DataFrame()
    selected_sub_bin = None
    
    # 準備兩個變數，用來記錄要顯示在右邊的 BIN_CODE 和 BIN
    display_bin_code = ""
    display_bin = ""

    # --- 步驟二：選擇 BIN CODE / BIN ---
    if search_method == "使用 BIN_CODE":
        unique_bin_codes = [x for x in base_df["BIN_CODE"].unique() if x != "無資料"]
        selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN_CODE", unique_bin_codes)
        
        if selected_val:
            filtered_df = base_df[base_df["BIN_CODE"] == selected_val]
            associated_bins = [x for x in filtered_df["BIN"].unique() if x != "無資料"]
            
            # 記錄要顯示在右邊的字串
            display_bin_code = selected_val
            display_bin = ', '.join(associated_bins) if associated_bins else "(無對應紀錄)"
            
            if associated_bins:
                st.caption(f"💡 對應的 BIN: {display_bin}")
            else:
                st.caption("💡 對應的 BIN: (無對應紀錄)")
            
    else:
        unique_bins = [x for x in base_df["BIN"].unique() if x != "無資料"]
        selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN", unique_bins)
        
        if selected_val:
            filtered_df = base_df[base_df["BIN"] == selected_val]
            associated_bin_codes = [x for x in filtered_df["BIN_CODE"].unique() if x != "無資料"]
            
            # 記錄要顯示在右邊的字串
            display_bin = selected_val
            display_bin_code = ', '.join(associated_bin_codes) if associated_bin_codes else "(無對應紀錄)"
            
            if associated_bin_codes:
                st.caption(f"💡 對應的 BIN_CODE: {display_bin_code}")
            else:
                st.caption("💡 對應的 BIN_CODE: (無對應紀錄)")

    st.divider()

    # --- 步驟三：選擇 SUB_BIN ---
    if not filtered_df.empty:
        unique_sub_bins = filtered_df["SUB_BIN"].unique()
        selected_sub_bin = st.selectbox("📑 第三步：請選擇 SUB_BIN", unique_sub_bins)


# ==========================================
# 🖥️ 主畫面區域 (顯示結果)
# ==========================================
st.title("🔍 故障排除查詢系統")

if filtered_df.empty or not selected_sub_bin:
    st.info("👈 請從左側選單 (手機請點擊左上角 〉圖示) 開始選擇條件。")
else:
    st.markdown("### 📌 查詢條件確認")
    
    # ★ 修改處：移除 col1, col2，改為直接上下顯示
    with st.container(border=True):
        st.markdown(f"**🔹 BIN_CODE:** {display_bin_code}")
        st.markdown(f"**🔸 BIN:** {display_bin}")
        
        # 顯示 SUB_BIN 全文
        if selected_sub_bin != "無資料":
            st.markdown(f"> **🏷️ SUB_BIN 全文：**  \n> {selected_sub_bin}")
            
    st.divider()
    
    # 進行最終過濾
    final_df = filtered_df[filtered_df["SUB_BIN"] == selected_sub_bin]
    
    st.subheader(f"💡 查詢結果 (共 {len(final_df)} 筆)")
    
    # 印出結果
    for index, row in final_df.iterrows():
        with st.container(border=True):
            st.markdown(f"**🚨 Possible Cause (可能原因):**\n{row['Possible Cause']}")
            st.markdown(f"**✅ Solution (解決方案):**\n{row['Solution']}")
            
            col1, col2 = st.columns(2)
            
            if row['Ref Log'] != "無資料":
                col1.caption(f"**Ref Log:** {row['Ref Log']}")
                
            if row['REV'] != "無資料":
                col2.caption(f"**REV:** {row['REV']}")
