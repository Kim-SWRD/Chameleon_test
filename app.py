import streamlit as st
import pandas as pd

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="centered")

# --- 自訂 CSS 樣式 ---
# 透過 CSS 將 Streamlit 預設的 primary 按鈕顏色改為綠色，避免紅色帶來錯誤的錯覺
st.markdown("""
<style>
button[kind="primary"] {
    background-color: #28a745 !important;
    border-color: #28a745 !important;
    color: white !important;
}
button[kind="primary"]:hover {
    background-color: #218838 !important;
    border-color: #1e7e34 !important;
}
</style>
""", unsafe_allow_html=True)

# --- 初始化 Session State ---
# 用於讓「按鈕點擊」與「輸入框」能夠互相連動
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
    # 讀取 mapping 資料，將全欄位轉為字串
    df = pd.read_excel("TS2_mapping.xlsx", dtype=str)
    # 過濾掉開頭的無效資料 (NO. 為 NaN 的列)
    df = df.dropna(subset=['NO.'])
    # 清理 NO. 欄位，避免 Pandas 將整數讀成 2.0，統一轉為整數字串
    df["NO."] = df["NO."].astype(str).str.replace(r'\.0$', '', regex=True)
    
    # 將 DataFrame 內的 NO. 欄位更名為 TS2#，配合下拉選單顯示
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

st.title("🛠️ 伺服器系統查詢工具")

# ==========================================
# 📑 建立頂部切換分頁
# ==========================================
tab_rca, tab_map = st.tabs(["🔍 RCA 故障排除查詢", "🔄 TS2 SN Mapping 查詢"])

# ==========================================
# 分頁 1: RCA 故障排除查詢
# ==========================================
with tab_rca:
    st.header("🔍 RCA 故障排除查詢")
    
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
    st.header("🔄 TS2 SN Mapping 查詢")
    
    # 定義按鈕的回呼函式：點擊按鈕後，自動將搜尋條件切換為 TS2# 並填入對應數字
    def set_ts2_search(num_str):
        st.session_state["map_col"] = "TS2#"
        st.session_state["map_search_input"] = num_str

    # 取得目前 df 中所有有效的 TS2# 集合，方便快速比對
    valid_ts2 = set(df_map["TS2#"].dropna().astype(str).tolist())
    
    # 統計 1~99 之間有幾筆資料
    valid_count = sum(1 for i in range(1, 100) if str(i) in valid_ts2)

    # 1. 建立 1~99 的快速點選面板 (動態顯示統計數量與綠色提示)
    panel_title = f"🎛️ TS2# 快速點選面板 (綠色: 有資料 ({valid_count}筆) / 灰色: 無資料)"
    with st.expander(panel_title, expanded=True):
        # 建立 10 列 x 10 欄的按鈕矩陣
        for row in range(10):
            cols = st.columns(10)
            for col_idx in range(10):
                num = row * 10 + col_idx + 1
                if num > 99:
                    break
                
                # 判斷該號碼是否有資料，有資料套用 primary (已被 CSS 改為綠色)，無資料套用 secondary (預設灰色)
                is_valid = str(num) in valid_ts2
                btn_type = "primary" if is_valid else "secondary"
                
                # 建立按鈕並綁定 on_click 事件，文字只顯示純數字
                cols[col_idx].button(
                    str(num), 
                    key=f"btn_quick_{num}", 
                    on_click=set_ts2_search, 
                    args=(str(num),),
                    type=btn_type,
                    use_container_width=True
                )

    # 2. 原本的搜尋區塊 (保留並支援與按鈕連動)
    with st.container(border=True):
        st.markdown("輸入 **TS2# NO.** (例如: 2), 或是輸入 **CSM BASE, CSM TRAY, FULL SYS** 任意一組 SN，即可互相反查。")
        
        map_cols = ["TS2#", "CSM BASE", "CSM TRAY", "FULL SYS"]
        
        col1, col2 = st.columns([1, 2])
        with col1:
            search_col = st.selectbox("📌 選擇查詢條件", map_cols, key="map_col")
        with col2:
            search_val = st.text_input(f"✍️ 請輸入 {search_col}", key="map_search_input").strip()
            
    # 3. 執行搜尋比對
    if search_val:
        # 智慧處理：自動過濾 "TS2#" 等字眼，只取數字
        if search_col == "TS2#":
            search_val = search_val.upper().replace("TS2#", "").replace("TS#", "").strip()
        
        # 進行條件比對
        match_df = df_map[df_map[search_col] == search_val]
        
        if not match_df.empty:
            st.success("✅ 找到對應的 SN 關聯資料！")
            
            for idx, row in match_df.iterrows():
                with st.container(border=True):
                    # 固定顯示 TS2# 前綴作為標題
                    st.markdown(f"### 🔹 系統標號：TS2#{row['TS2#']}")
                    
                    # 使用三個 Column 並列顯示對應的三個重要 SN
                    c1, c2, c3 = st.columns(3)
                    c1.info(f"**CSM BASE**  \n{row['CSM BASE']}")
                    c2.info(f"**CSM TRAY**  \n{row['CSM TRAY']}")
                    c3.info(f"**FULL SYS**  \n{row['FULL SYS']}")
                    
                    # 補充顯示旁邊的測試狀態
                    st.caption(f"🔍 站點狀態 👉 JTAG: `{row.get('JTAG', '無資料')}` | AOT: `{row.get('AOT', '無資料')}` | FT: `{row.get('FT', '無資料')}`")
        else:
            st.error(f"⚠️ 找不到 {search_col} = `{search_val}` 的資料，請確認輸入是否有誤。")