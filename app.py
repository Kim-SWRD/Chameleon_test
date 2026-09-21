import streamlit as st
import pandas as pd

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="centered")

# --- 自訂 CSS 樣式 ---
st.markdown("""
<style>
/* 1. 將預設 primary 按鈕改為綠色 */
button[kind="primary"] {
    background-color: #28a745 !important;
    border-color: #28a745 !important;
    color: white !important;
}
button[kind="primary"]:hover {
    background-color: #218838 !important;
    border-color: #1e7e34 !important;
}

/* 2. 放大 Tab 分頁的標題字體 */
button[data-baseweb="tab"] p {
    font-size: 20px !important;
    font-weight: 700 !important;
}

/* 3. 強制 st.columns 在所有螢幕尺寸下 (包含手機直向) 保持水平排列 */
/* Streamlit 1.30+ 版本使用 data-testid="stHorizontalBlock" */
[data-testid="stHorizontalBlock"] {
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    overflow-x: auto !important; /* 如果螢幕真的太小，允許橫向捲動，避免按鈕被壓扁到看不見 */
}

/* 強制每個 column 佔據相等的寬度 */
[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
    width: calc(10% - 0.2rem) !important;
    flex: 1 1 auto !important;
    min-width: 0 !important;
    padding: 0 0.1rem !important; /* 縮小按鈕之間的間距 */
}

/* 縮小按鈕內的 padding 讓它在手機上能擠得下 */
[data-testid="stHorizontalBlock"] > div[data-testid="column"] button {
    padding-left: 0 !important;
    padding-right: 0 !important;
    width: 100% !important;
    font-size: 12px !important; 
}
</style>
""", unsafe_allow_html=True)

# --- 初始化 Session State ---
if "map_col" not in st.session_state:
    st.session_state["map_col"] = "TS2#"
if "map_search_input" not in st.session_state:
    st.session_state["map_search_input"] = ""

# --- 讀取與快取資料 ---
@st.cache_data(ttl=60)
def load_rca_data():
    df = pd.read_excel("RCA.xlsx", dtype=str)
    df.fillna("無資料", inplace=True)
    return df

@st.cache_data(ttl=60)
def load_mapping_data():
    df = pd.read_excel("TS2_mapping.xlsx", dtype=str)
    df = df.dropna(subset=['NO.'])
    df["NO."] = df["NO."].astype(str).str.replace(r'\.0$', '', regex=True)
    df.rename(columns={"NO.": "TS2#"}, inplace=True) 
    df.fillna("無資料", inplace=True)
    return df

try:
    df_rca = load_rca_data()
except FileNotFoundError:
    st.error("找不到 RCA.xlsx 檔案！請確認它是否放在跟 app.py 同一個資料夾內。")
    st.stop()

try:
    df_map = load_mapping_data()
except FileNotFoundError:
    st.error("找不到 TS2_mapping.xlsx 檔案！請確認它是否放在跟 app.py 同一個資料夾內。")
    st.stop()

# ==========================================
# 📑 建立頂部切換分頁
# ==========================================
tab_rca, tab_map = st.tabs(["🔍 RCA 故障排除查詢", "🔄 TS2 SN Mapping 查詢"])

# ==========================================
# 分頁 1: RCA 故障排除查詢
# ==========================================
with tab_rca:
    with st.container(border=True):
        unique_stations = [x for x in df_rca["STATION"].unique() if x != "無資料"]
        station_options = ["ALL"] + unique_stations
        
        selected_station = st.selectbox("📌 選擇站別", station_options, key="rca_station")
        
        if selected_station == "ALL":
            base_df = df_rca
        else:
            base_df = df_rca[df_rca["STATION"] == selected_station]

        search_method = st.radio("🔍 第一步：選擇查詢方式", ["使用 BIN_CODE", "使用 BIN"], horizontal=True, key="rca_method")

        filtered_df = pd.DataFrame()
        selected_sub_bin = None
        
        display_bin_code = ""
        display_bin = ""

        if search_method == "使用 BIN_CODE":
            unique_bin_codes = [x for x in base_df["BIN_CODE"].unique() if x != "無資料"]
            selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN_CODE", unique_bin_codes, key="rca_bincode")
            
            if selected_val:
                filtered_df = base_df[base_df["BIN_CODE"] == selected_val]
                associated_bins = [x for x in filtered_df["BIN"].unique() if x != "無資料"]
                
                display_bin_code = selected_val
                display_bin = ', '.join(associated_bins) if associated_bins else "(無對應紀錄)"
                
                st.caption(f"💡 對應 BIN: {display_bin}")
                
        else:
            unique_bins = [x for x in base_df["BIN"].unique() if x != "無資料"]
            selected_val = st.selectbox("🏷️ 第二步：請選擇 BIN", unique_bins, key="rca_bin")
            
            if selected_val:
                filtered_df = base_df[base_df["BIN"] == selected_val]
                associated_bin_codes = [x for x in filtered_df["BIN_CODE"].unique() if x != "無資料"]
                
                display_bin = selected_val
                display_bin_code = ', '.join(associated_bin_codes) if associated_bin_codes else "(無對應紀錄)"
                
                st.caption(f"💡 對應 BIN_CODE: {display_bin_code}")
                st.caption(f"💡 BIN 全文: {display_bin}")

        if not filtered_df.empty:
            unique_sub_bins = filtered_df["SUB_BIN"].unique()
            selected_sub_bin = st.selectbox("📑 第三步：請選擇 SUB_BIN", unique_sub_bins, key="rca_subbin")

    # 顯示 RCA 結果區域
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


# ==========================================
# 分頁 2: TS2 SN Mapping 查詢
# ==========================================
with tab_map:
    def set_ts2_search(num_str):
        st.session_state["map_col"] = "TS2#"
        st.session_state["map_search_input"] = num_str

    valid_ts2 = set(df_map["TS2#"].dropna().astype(str).tolist())
    valid_count = sum(1 for i in range(1, 100) if str(i) in valid_ts2)

    panel_title = f"🎛️ TS2# 快速點選面板 (綠色: 有資料 ({valid_count}筆) / 灰色: 無資料)"
    with st.expander(panel_title, expanded=True):
        # 建立 10 列 x 10 欄的按鈕矩陣
        for row in range(10):
            # 這裡的 st.columns(10) 會被上方的 CSS 強制設定為 nowrap 且 flex-direction: row
            cols = st.columns(10)
            for col_idx in range(10):
                num = row * 10 + col_idx + 1
                if num > 99:
                    break
                
                is_valid = str(num) in valid_ts2
                btn_type = "primary" if is_valid else "secondary"
                
                cols[col_idx].button(
                    str(num), 
                    key=f"btn_quick_{num}", 
                    on_click=set_ts2_search, 
                    args=(str(num),),
                    type=btn_type,
                    use_container_width=True
                )

    with st.container(border=True):
        st.markdown("輸入 **TS2# NO.** (例如: 2), 或是輸入 **CSM BASE, CSM TRAY, FULL SYS** 任意一組 SN，即可互相反查。")
        
        map_cols = ["TS2#", "CSM BASE", "CSM TRAY", "FULL SYS"]
        
        col1, col2 = st.columns([1, 2])
        with col1:
            search_col = st.selectbox("📌 選擇查詢條件", map_cols, key="map_col")
        with col2:
            search_val = st.text_input(f"✍️ 請輸入 {search_col}", key="map_search_input").strip()
            
    if search_val:
        if search_col == "TS2#":
            search_val = search_val.upper().replace("TS2#", "").replace("TS#", "").strip()
        
        match_df = df_map[df_map[search_col] == search_val]
        
        if not match_df.empty:
            st.success("✅ 找到對應的 SN 關聯資料！")
            
            for idx, row in match_df.iterrows():
                with st.container(border=True):
                    st.markdown("### 🔹 系統標號")
                    st.code(f"TS2#{row['TS2#']}", language="plaintext")
                    
                    # 這裡的 c1, c2, c3 也是 st.columns，為了不影響它們在手機上自然往下疊的行為
                    # 我們只在 CSS 裡鎖定了 "stHorizontalBlock"，但這可能會一併影響這三個欄位。
                    # 如果你發現這三個欄位在手機上也變成水平排列且擠在一起，
                    # 可以在 CSS 中替換選擇器，或者就讓它們水平排列，因為字串可以用 st.code 滾動顯示。
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.markdown("**CSM BASE**")
                        if pd.notna(row['CSM BASE']) and row['CSM BASE'] != "無資料":
                            st.code(row['CSM BASE'], language="plaintext")
                        else:
                            st.info("無資料")
                            
                    with c2:
                        st.markdown("**CSM TRAY**")
                        if pd.notna(row['CSM TRAY']) and row['CSM TRAY'] != "無資料":
                            st.code(row['CSM TRAY'], language="plaintext")
                        else:
                            st.info("無資料")
                            
                    with c3:
                        st.markdown("**FULL SYS**")
                        if pd.notna(row['FULL SYS']) and row['FULL SYS'] != "無資料":
                            st.code(row['FULL SYS'], language="plaintext")
                        else:
                            st.info("無資料")
                    
                    st.caption(f"🔍 站點狀態 👉 JTAG: `{row.get('JTAG', '無資料')}` | AOT: `{row.get('AOT', '無資料')}` | FT: `{row.get('FT', '無資料')}`")
        else:
            st.error(f"⚠️ 找不到 {search_col} = `{search_val}` 的資料，請確認輸入是否有誤。")