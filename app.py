import streamlit as st
import pandas as pd
import io
from github import Github

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="centered")

# --- 自訂 CSS 樣式 ---
st.markdown("""
<style>
button[kind="primary"] { background-color: #28a745 !important; border-color: #28a745 !important; color: white !important; }
button[kind="primary"]:hover { background-color: #218838 !important; border-color: #1e7e34 !important; }
button[kind="tertiary"] { background-color: #ff9800 !important; border-color: #ff9800 !important; color: white !important; }
button[kind="tertiary"]:hover { background-color: #e68a00 !important; border-color: #e68a00 !important; }
button[data-baseweb="tab"] p { font-size: 20px !important; font-weight: 700 !important; }

/* CSS Grid 強制 10 欄網格佈局 */
div[data-testid="stExpanderDetails"] div[data-testid="stHorizontalBlock"] {
    display: grid !important; grid-template-columns: repeat(10, 1fr) !important; gap: 4px !important; width: 100% !important; padding-bottom: 3px !important;
}
div[data-testid="stExpanderDetails"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"] { width: 100% !important; min-width: 0 !important; padding: 0 !important; }
div[data-testid="stExpanderDetails"] div[data-testid="stHorizontalBlock"] button {
    width: 200% !important; aspect-ratio: 1 / 1 !important; border-radius: 6px !important; padding: 0 !important; margin: 0 !important; min-height: 0 !important; height: auto !important; display: flex !important; align-items: center !important; justify-content: center !important;
}
div[data-testid="stExpanderDetails"] div[data-testid="stHorizontalBlock"] button > div { display: flex !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; margin: 0 !important; padding: 0 !important; }
div[data-testid="stExpanderDetails"] div[data-testid="stHorizontalBlock"] button p { font-size: 13px !important; font-weight: 700 !important; margin: 0 !important; padding: 0 !important; text-align: center !important; }
div[data-testid="stCodeBlock"] button { opacity: 1 !important; visibility: visible !important; display: inline-flex !important; }
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
# 分頁 2: TS2 SN Mapping 查詢 (含 GitHub 上傳功能)
# ==========================================
with tab_map:
    def set_ts2_search(num_str):
        st.session_state["map_col"] = "TS2#"
        st.session_state["map_search_input"] = num_str

    valid_ts2 = set(df_map["TS2#"].dropna().astype(str).tolist())
    valid_count = sum(1 for i in range(1, 100) if str(i) in valid_ts2)
    
    current_search_col = st.session_state.get("map_col", "TS2#")
    current_search_val = st.session_state.get("map_search_input", "").strip()
    active_ts2_numbers = set()
    
    if current_search_val:
        query_val = current_search_val
        if current_search_col == "TS2#":
            query_val = query_val.upper().replace("TS2#", "").replace("TS#", "").strip()
        temp_df = df_map[df_map[current_search_col] == query_val]
        active_ts2_numbers = set(temp_df["TS2#"].dropna().astype(str).tolist())

    panel_title = f"🎛️ TS2# 快速點選面板 (綠色: 有資料 ({valid_count}筆) / 橘色: 目前選取)"
    with st.expander(panel_title, expanded=True):
        for row in range(10):
            cols = st.columns(10)
            for col_idx in range(10):
                num = row * 10 + col_idx + 1
                if num > 99:
                    break
                
                is_valid = str(num) in valid_ts2
                is_selected = str(num) in active_ts2_numbers
                
                if is_selected:
                    btn_type = "tertiary"
                elif is_valid:
                    btn_type = "primary"
                else:
                    btn_type = "secondary"
                
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
                # 紀錄當前的 TS2#
                ts2_id = row['TS2#']
                
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.7, 0.3], vertical_alignment="center")
                    with c_title:
                        st.markdown(f"### 🔹 系統標號：TS2#{ts2_id}")
                    with c_toggle:
                        is_editing = st.toggle("✏️ 進入編輯模式", key=f"toggle_{ts2_id}")
                    
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.markdown("**CSM BASE**")
                        val_base = row['CSM BASE'] if pd.notna(row['CSM BASE']) and row['CSM BASE'] != "無資料" else ""
                        if is_editing:
                            st.text_input("CSM BASE", value=val_base, label_visibility="collapsed", key=f"edit_base_{ts2_id}")
                        else:
                            st.code(val_base if val_base else "無資料", language="plaintext")
                            
                    with c2:
                        st.markdown("**CSM TRAY**")
                        val_tray = row['CSM TRAY'] if pd.notna(row['CSM TRAY']) and row['CSM TRAY'] != "無資料" else ""
                        if is_editing:
                            st.text_input("CSM TRAY", value=val_tray, label_visibility="collapsed", key=f"edit_tray_{ts2_id}")
                        else:
                            st.code(val_tray if val_tray else "無資料", language="plaintext")
                            
                    with c3:
                        st.markdown("**FULL SYS**")
                        val_full = row['FULL SYS'] if pd.notna(row['FULL SYS']) and row['FULL SYS'] != "無資料" else ""
                        if is_editing:
                            st.text_input("FULL SYS", value=val_full, label_visibility="collapsed", key=f"edit_full_{ts2_id}")
                        else:
                            st.code(val_full if val_full else "無資料", language="plaintext")
                    
                    if is_editing:
                        st.divider()
                        st.markdown("#### 🔍 站點狀態")
                        s_c1, s_c2, s_c3 = st.columns(3)
                        status_opts = ["無資料", "PASS", "FAIL"]
                        
                        def get_opt_idx(val):
                            if pd.isna(val) or str(val).strip() == "無資料": return 0
                            v = str(val).strip().upper()
                            if v == "PASS": return 1
                            if v == "FAIL": return 2
                            return 0

                        with s_c1:
                            st.selectbox("JTAG", status_opts, index=get_opt_idx(row.get('JTAG')), key=f"edit_jtag_{ts2_id}")
                        with s_c2:
                            st.selectbox("AOT", status_opts, index=get_opt_idx(row.get('AOT')), key=f"edit_aot_{ts2_id}")
                        with s_c3:
                            st.selectbox("FT", status_opts, index=get_opt_idx(row.get('FT')), key=f"edit_ft_{ts2_id}")
                            
                        st.write("") 
                        
                        # ★ 實際的儲存與上傳邏輯 ★
                        if st.button("💾 儲存修改並同步至 GitHub", type="primary", use_container_width=True):
                            # 1. 檢查是否有設定 Secrets
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！請先在 .streamlit/secrets.toml 中設定。")
                            else:
                                with st.spinner("🔄 正在更新並上傳至 GitHub..."):
                                    try:
                                        # 2. 取得使用者修改後的新值
                                        new_base = st.session_state.get(f"edit_base_{ts2_id}", "").strip()
                                        new_tray = st.session_state.get(f"edit_tray_{ts2_id}", "").strip()
                                        new_full = st.session_state.get(f"edit_full_{ts2_id}", "").strip()
                                        new_jtag = st.session_state.get(f"edit_jtag_{ts2_id}", "無資料")
                                        new_aot = st.session_state.get(f"edit_aot_{ts2_id}", "無資料")
                                        new_ft = st.session_state.get(f"edit_ft_{ts2_id}", "無資料")

                                        # 3. 更新 df_map 中的對應列 (因為讀取時欄位名改成了 TS2#，原寫入前要確定對應)
                                        idx_update = df_map[df_map['TS2#'] == ts2_id].index
                                        df_map.loc[idx_update, 'CSM BASE'] = new_base if new_base else "無資料"
                                        df_map.loc[idx_update, 'CSM TRAY'] = new_tray if new_tray else "無資料"
                                        df_map.loc[idx_update, 'FULL SYS'] = new_full if new_full else "無資料"
                                        df_map.loc[idx_update, 'JTAG'] = new_jtag
                                        df_map.loc[idx_update, 'AOT'] = new_aot
                                        df_map.loc[idx_update, 'FT'] = new_ft

                                        # 準備要覆蓋回去的 df (需要把 TS2# 改回原本的 NO. 避免破壞 Excel 原有欄位名稱)
                                        df_upload = df_map.copy()
                                        df_upload.rename(columns={"TS2#": "NO."}, inplace=True)
                                        # 如果 Excel 檔裡有把 "無資料" 當成空值，可以在這裡把 "無資料" 取代為 None (可選)
                                        
                                        # 4. 轉換為 Excel 的二進位資料
                                        output = io.BytesIO()
                                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                            df_upload.to_excel(writer, index=False)
                                        excel_data = output.getvalue()
                                        
                                        # 5. 連線 GitHub API 進行檔案覆蓋
                                        g = Github(st.secrets["GITHUB_TOKEN"])
                                        repo = g.get_repo(st.secrets["GITHUB_REPO"])
                                        
                                        file_path = "TS2_mapping.xlsx" # 你的檔案在 repo 中的確切路徑
                                        contents = repo.get_contents(file_path)
                                        
                                        repo.update_file(
                                            path=contents.path,
                                            message=f"Update TS2#{ts2_id} via Streamlit",
                                            content=excel_data,
                                            sha=contents.sha
                                        )
                                        
                                        # 6. 清除快取，讓畫面重整後抓到最新資料
                                        st.cache_data.clear()
                                        st.success("✅ 成功同步至 GitHub！檔案已更新。")
                                        
                                    except Exception as e:
                                        st.error(f"❌ 上傳失敗: {e}")
                    else:
                        st.caption(f"🔍 站點狀態 👉 JTAG: `{row.get('JTAG', '無資料')}` | AOT: `{row.get('AOT', '無資料')}` | FT: `{row.get('FT', '無資料')}`")
        else:
            st.error(f"⚠️ 找不到 {search_col} = `{search_val}` 的資料，請確認輸入是否有誤。")