import streamlit as st
import pandas as pd

# 設定網頁標題與寬度
st.set_page_config(page_title="卡美問題查詢系統", layout="centered")

st.title("🔍 故障排除查詢系統")
st.write("請根據下方的條件篩選出對應的解決方案。")

# --- 1. 讀取與快取資料 ---
@st.cache_data
def load_data():
    df = pd.read_excel("RCA.xlsx", dtype=str)
    df.fillna("無資料", inplace=True)
    return df

try:
    df = load_data()
except FileNotFoundError:
    st.error("找不到 RCA.xlsx 檔案！請確認它是否放在跟 app.py 同一個資料夾 (工作目錄) 內。")
    st.stop()

# --- ★ STATION 篩選 (過濾器) ---
st.divider()
st.subheader("🛠️ 選擇站別 (STATION)")

# 抓出不重複的 STATION，剔除無資料，並在最前面加上 "ALL"
unique_stations = [x for x in df["STATION"].unique() if x != "無資料"]
station_options = ["ALL"] + unique_stations

# horizontal=True 讓選項像按鈕一樣水平並排
selected_station = st.radio("請選擇要過濾的站別 (預設為 ALL)：", station_options, horizontal=True)

# 根據選到的 STATION 產生一份「基礎資料庫 (base_df)」
if selected_station == "ALL":
    base_df = df
else:
    base_df = df[df["STATION"] == selected_station]

# --- 2. 選擇查詢方式 ---
st.divider()
search_method = st.radio("第一步：請選擇查詢方式", ["使用 BIN_CODE", "使用 BIN"])

filtered_df = pd.DataFrame()

# --- 3. 根據選擇顯示第一個下拉選單，並顯示對應資訊 ---
if search_method == "使用 BIN_CODE":
    unique_bin_codes = [x for x in base_df["BIN_CODE"].unique() if x != "無資料"]
    selected_val = st.selectbox("第二步：請選擇 BIN_CODE", unique_bin_codes)
    
    if selected_val:
        filtered_df = base_df[base_df["BIN_CODE"] == selected_val]
        associated_bins = [x for x in filtered_df["BIN"].unique() if x != "無資料"]
        if associated_bins:
            st.info(f"💡 對應的 BIN 為： {', '.join(associated_bins)}")
        else:
            st.info("💡 對應的 BIN 為： (無對應紀錄)")
        
else:
    unique_bins = [x for x in base_df["BIN"].unique() if x != "無資料"]
    selected_val = st.selectbox("第二步：請選擇 BIN", unique_bins)
    
    if selected_val:
        filtered_df = base_df[base_df["BIN"] == selected_val]
        associated_bin_codes = [x for x in filtered_df["BIN_CODE"].unique() if x != "無資料"]
        if associated_bin_codes:
            st.info(f"💡 對應的 BIN_CODE 為： {', '.join(associated_bin_codes)}")
        else:
            st.info("💡 對應的 BIN_CODE 為： (無對應紀錄)")

# --- 4. 顯示 SUB_BIN 下拉選單 ---
if not filtered_df.empty:
    unique_sub_bins = filtered_df["SUB_BIN"].unique()
    selected_sub_bin = st.selectbox("第三步：請選擇 SUB_BIN", unique_sub_bins)
    
    # 在選單正下方，把 SUB_BIN 的全文完整印出來
    if selected_sub_bin and selected_sub_bin != "無資料":
        st.markdown(f"> **🏷️ SUB_BIN 全文：**  \n> {selected_sub_bin}")
    
    # 進行最終過濾
    final_df = filtered_df[filtered_df["SUB_BIN"] == selected_sub_bin]
    
    # --- 5. 顯示結果 (Possible Cause & Solution) ---
    st.subheader("💡 查詢結果")
    st.write(f"共找到 {len(final_df)} 筆對應的解決方案：")
    
    for index, row in final_df.iterrows():
        with st.container(border=True):
            st.markdown(f"**🚨 Possible Cause (可能原因):**\n{row['Possible Cause']}")
            st.markdown(f"**✅ Solution (解決方案):**\n{row['Solution']}")
            
            # 顯示參考 Log
            if row['Ref Log'] != "無資料":
                st.caption(f"Ref Log: {row['Ref Log']}")
                
            # ★ 新增：顯示 REV 欄位
            if row['REV'] != "無資料":
                st.caption(f"REV: {row['REV']}")