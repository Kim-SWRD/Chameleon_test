import streamlit as st
import pandas as pd
import io
import time
from github import Github

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="centered")

# --- 自訂 CSS 樣式 ---
st.markdown("""
<style>
button[kind="primary"] { background-color: #28a745 !important; border-color: #28a745 !important; color: white !important; }
button[kind="primary"]:hover { background-color: #218838 !important; border-color: #1e7e34 !important; }
button[kind="tertiary"] { background-color: #007bff !important; border-color: #007bff !important; color: white !important; }
button[kind="tertiary"]:hover { background-color: #0056b3 !important; border-color: #0056b3 !important; }
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
if "map_col" not in st.session_state: st.session_state["map_col"] = "TS2#"
if "map_search_input" not in st.session_state: st.session_state["map_search_input"] = ""
if "status_active_ts2" not in st.session_state: st.session_state["status_active_ts2"] = ""

# --- 讀取與快取資料 ---
@st.cache_data(ttl=60)
def load_rca_data():
    df = pd.read_excel("RCA.xlsx", dtype=str)
    df.fillna("無資料", inplace=True)
    return df

@st.cache_data(ttl=60)
def load_mapping_data():
    df = pd.read_excel("TS2_mapping.xlsx", dtype=str)
    rename_dict = {}
    for col in df.columns:
        col_upper = str(col).upper()
        if "NO." in col_upper or "TS2#" in col_upper: rename_dict[col] = "NO."
        elif "CSM BASE" in col_upper or "CSM_BASE" in col_upper: rename_dict[col] = "CSM BASE"
        elif "CSM TRAY" in col_upper or "CSM_TRAY" in col_upper: rename_dict[col] = "CSM TRAY"
        elif "FULL SYS" in col_upper or "FULL_SYS" in col_upper: rename_dict[col] = "FULL SYS"
        elif "JTAG" in col_upper: rename_dict[col] = "JTAG"
        elif "AOT" in col_upper: rename_dict[col] = "AOT"
        elif "FT" in col_upper: rename_dict[col] = "FT"
        elif "STATUS" in col_upper: rename_dict[col] = "STATUS"
        elif "OWNER" in col_upper: rename_dict[col] = "OWNER"
        elif "FAILURE BIN" in col_upper or "FAIL BIN" in col_upper or "NOTE" in col_upper: rename_dict[col] = "Failure BIN"
            
    df.rename(columns=rename_dict, inplace=True)
    if 'NO.' not in df.columns:
        st.error("Excel 中找不到包含 'NO.' 的欄位，請檢查檔案標題列是否有誤。")
        st.stop()
        
    df['NO.'] = df['NO.'].ffill()
    df = df.dropna(subset=['NO.'])
    df = df.groupby('NO.', as_index=False).first()
    df["NO."] = df["NO."].astype(str).str.replace(r'\.0$', '', regex=True)
    df.rename(columns={"NO.": "TS2#"}, inplace=True) 
    df.fillna("無資料", inplace=True)
    return df

@st.cache_data(ttl=60)
def load_note_data():
    df = pd.read_excel("TS2_note.xlsx", dtype=str)
    if 'NO.' in df.columns:
        df['NO.'] = df['NO.'].astype(str).str.replace(r'\.0$', '', regex=True)
    if 'NOTICE' not in df.columns:
        df['NOTICE'] = '0'
    df.fillna("無資料", inplace=True)
    return df

try: df_rca = load_rca_data()
except FileNotFoundError: st.error("找不到 RCA.xlsx 檔案！請確認它是否與 app.py 放在一起。"); st.stop()

try: df_map = load_mapping_data()
except FileNotFoundError: st.error("找不到 TS2_mapping.xlsx 檔案！請確認它是否與 app.py 放在一起。"); st.stop()

try: df_note = load_note_data()
except FileNotFoundError: st.error("找不到 TS2_note.xlsx 檔案！請確認它是否與 app.py 放在一起。"); st.stop()

station_opts = ["無資料", "PASS", "FAIL"]
def get_station_idx(val):
    if pd.isna(val) or str(val).strip() == "無資料": return 0
    v = str(val).strip().upper()
    if v == "PASS": return 1
    if v == "FAIL": return 2
    return 0

# ==========================================
# 📑 建立頂部切換分頁
# ==========================================
tab_rca, tab_map, tab_status = st.tabs(["🔍 故障排除", "🔄 Mapping查詢", "📊 TS2 STATUS"])

# ==========================================
# 分頁 1: 故障排除
# ==========================================
with tab_rca:
    st.header("🔍 故障排除")
    with st.container(border=True):
        unique_stations = [x for x in df_rca["STATION"].unique() if x != "無資料"]
        station_options = ["ALL"] + unique_stations
        selected_station = st.selectbox("📌 選擇站別", station_options, key="rca_station")
        base_df = df_rca if selected_station == "ALL" else df_rca[df_rca["STATION"] == selected_station]

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
                if row['Ref Log'] != "無資料": meta_info.append(f"**Log:** {row['Ref Log']}")
                if row['REV'] != "無資料": meta_info.append(f"**REV:** {row['REV']}")
                if meta_info: st.caption(" | ".join(meta_info))


# ==========================================
# 分頁 2: Mapping查詢
# ==========================================
with tab_map:
    col_title, col_upload = st.columns([0.6, 0.4])
    with col_title:
        st.header("🔄 Mapping查詢")
    with col_upload:
        with st.expander("📤 上傳 / 下載 TS2_mapping", expanded=False):
            try:
                with open("TS2_mapping.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前 Mapping 檔", data=f, file_name="TS2_mapping.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError:
                pass
                
            st.divider()
            uploaded_file = st.file_uploader("選擇 Mapping 檔案 (.xlsx)", type=["xlsx"])
            if uploaded_file:
                try:
                    df_test = pd.read_excel(uploaded_file, dtype=str)
                    rename_test = {}
                    for col in df_test.columns:
                        c_up = str(col).upper()
                        if "NO." in c_up or "TS2#" in c_up: rename_test[col] = "NO."
                        elif "CSM BASE" in c_up or "CSM_BASE" in c_up: rename_test[col] = "CSM BASE"
                        elif "CSM TRAY" in c_up or "CSM_TRAY" in c_up: rename_test[col] = "CSM TRAY"
                        elif "FULL SYS" in c_up or "FULL_SYS" in c_up: rename_test[col] = "FULL SYS"
                    df_test.rename(columns=rename_test, inplace=True)
                    
                    if 'NO.' not in df_test.columns:
                        st.error("❌ 嚴重錯誤：找不到 'NO.' 或 'TS2#' 欄位，無法解析此檔案。")
                    else:
                        missing_sn = [req_col for req_col in ['CSM BASE', 'CSM TRAY', 'FULL SYS'] if req_col not in df_test.columns]
                        st.success(f"✅ 驗證通過！共讀取到 {len(df_test)} 筆資料。")
                        if missing_sn: st.warning(f"⚠️ 警告：檔案缺少以下欄位 ({', '.join(missing_sn)})，仍可強制上傳。")
                            
                        if st.button("🚀 確認上傳並覆蓋至 GitHub", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 上傳中..."):
                                    uploaded_file.seek(0)
                                    excel_bytes = uploaded_file.read()
                                    with open("TS2_mapping.xlsx", "wb") as f: f.write(excel_bytes)
                                    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                    contents = repo.get_contents("TS2_mapping.xlsx")
                                    repo.update_file(contents.path, "Update TS2_mapping.xlsx via Streamlit Upload", excel_bytes, contents.sha)
                                    st.cache_data.clear()
                                    st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                    time.sleep(1.5)
                                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")

    def set_ts2_search(num_str):
        st.session_state["map_col"] = "TS2#"
        st.session_state["map_search_input"] = num_str

    def get_sn_count(row):
        cnt = 0
        for col in ['CSM BASE', 'CSM TRAY', 'FULL SYS']:
            val = row.get(col, "無資料")
            if pd.notna(val) and str(val).strip() not in ["", "無資料", "nan", "NaN"]: cnt += 1
        return cnt

    ts2_sn_counts = {}
    for idx, row in df_map.iterrows():
        ts2_id = str(row['TS2#']).strip()
        cnt = get_sn_count(row)
        if ts2_id not in ts2_sn_counts or cnt > ts2_sn_counts[ts2_id]: ts2_sn_counts[ts2_id] = cnt

    current_search_col = st.session_state.get("map_col", "TS2#")
    current_search_val = st.session_state.get("map_search_input", "").strip()
    active_ts2_numbers = set()
    
    if current_search_val:
        query_val = current_search_val
        if current_search_col == "TS2#": query_val = query_val.upper().replace("TS2#", "").replace("TS#", "").strip()
        temp_df = df_map[df_map[current_search_col] == query_val]
        active_ts2_numbers = set(temp_df["TS2#"].dropna().astype(str).tolist())

    dynamic_yellow_css_t2 = ""
    full_cnt, partial_cnt, empty_cnt = 0, 0, 0
    
    for num in range(1, 100):
        c = ts2_sn_counts.get(str(num), 0)
        is_selected = str(num) in active_ts2_numbers
        
        if c == 3: full_cnt += 1
        elif c in [1, 2]: partial_cnt += 1
        else: empty_cnt += 1
        
        if not is_selected and c in [1, 2]:
            dynamic_yellow_css_t2 += f"""
            div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }}
            div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}
            """
    
    if dynamic_yellow_css_t2: st.markdown(f"<style>{dynamic_yellow_css_t2}</style>", unsafe_allow_html=True)

    panel_title_t2 = f"🎛️ TS2# 快速點選面板 (綠色: 完整({full_cnt}) / 黃色: 缺件({partial_cnt}) / 灰色: 無資料({empty_cnt}) / 藍色: 選取)"
    
    with st.expander(panel_title_t2, expanded=True):
        st.markdown('<div class="t2-panel" style="display:none;"></div>', unsafe_allow_html=True)
        cols = st.columns(99)
        for num in range(1, 100):
            c = ts2_sn_counts.get(str(num), 0)
            is_selected = str(num) in active_ts2_numbers
            
            if is_selected: btn_type = "tertiary"
            elif c == 3: btn_type = "primary"
            else: btn_type = "secondary"
                
            cols[num-1].button(str(num), key=f"btn_t2_{num}", on_click=set_ts2_search, args=(str(num),), type=btn_type, use_container_width=True)

    with st.container(border=True):
        st.markdown("輸入 **TS2# NO.** (例如: 2), 或是輸入 **CSM BASE, CSM TRAY, FULL SYS** 任意一組 SN，即可互相反查。")
        map_cols = ["TS2#", "CSM BASE", "CSM TRAY", "FULL SYS"]
        col1, col2 = st.columns([1, 2])
        with col1: search_col = st.selectbox("📌 選擇查詢條件", map_cols, key="map_col")
        with col2: search_val = st.text_input(f"✍️ 請輸入 {search_col}", key="map_search_input").strip()
            
    if search_val:
        if search_col == "TS2#": search_val = search_val.upper().replace("TS2#", "").replace("TS#", "").strip()
        match_df = df_map[df_map[search_col] == search_val]
        
        if not match_df.empty:
            st.success("✅ 找到對應的 SN 關聯資料！")
            for idx, row in match_df.iterrows():
                ts2_id = row['TS2#']
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.7, 0.3], vertical_alignment="center")
                    with c_title: st.markdown(f"### 🔹 系統標號：TS2#{ts2_id}")
                    with c_toggle: is_editing = st.toggle("✏️ 進入編輯模式", key=f"t2_toggle_{ts2_id}")
                    
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.markdown("**CSM BASE**")
                        val_base = row['CSM BASE'] if pd.notna(row['CSM BASE']) and row['CSM BASE'] != "無資料" else ""
                        if is_editing: st.text_input("CSM BASE", value=val_base, label_visibility="collapsed", key=f"t2_edit_base_{ts2_id}")
                        else: st.code(val_base if val_base else "無資料", language="plaintext")
                            
                    with c2:
                        st.markdown("**CSM TRAY**")
                        val_tray = row['CSM TRAY'] if pd.notna(row['CSM TRAY']) and row['CSM TRAY'] != "無資料" else ""
                        if is_editing: st.text_input("CSM TRAY", value=val_tray, label_visibility="collapsed", key=f"t2_edit_tray_{ts2_id}")
                        else: st.code(val_tray if val_tray else "無資料", language="plaintext")
                            
                    with c3:
                        st.markdown("**FULL SYS**")
                        val_full = row['FULL SYS'] if pd.notna(row['FULL SYS']) and row['FULL SYS'] != "無資料" else ""
                        if is_editing: st.text_input("FULL SYS", value=val_full, label_visibility="collapsed", key=f"t2_edit_full_{ts2_id}")
                        else: st.code(val_full if val_full else "無資料", language="plaintext")
                    
                    if is_editing:
                        st.divider()
                        st.markdown("#### 🔍 站點狀態")
                        s_c1, s_c2, s_c3 = st.columns(3)
                        with s_c1: st.selectbox("JTAG", station_opts, index=get_station_idx(row.get('JTAG')), key=f"t2_edit_jtag_{ts2_id}")
                        with s_c2: st.selectbox("AOT", station_opts, index=get_station_idx(row.get('AOT')), key=f"t2_edit_aot_{ts2_id}")
                        with s_c3: st.selectbox("FT", station_opts, index=get_station_idx(row.get('FT')), key=f"t2_edit_ft_{ts2_id}")
                            
                        st.write("") 
                        if st.button("💾 儲存修改並同步至 GitHub", key=f"t2_save_btn_{ts2_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新並上傳..."):
                                    try:
                                        idx_update = df_map[df_map['TS2#'] == ts2_id].index
                                        df_map.loc[idx_update, 'CSM BASE'] = st.session_state.get(f"t2_edit_base_{ts2_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'CSM TRAY'] = st.session_state.get(f"t2_edit_tray_{ts2_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'FULL SYS'] = st.session_state.get(f"t2_edit_full_{ts2_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'JTAG'] = st.session_state.get(f"t2_edit_jtag_{ts2_id}", "無資料")
                                        df_map.loc[idx_update, 'AOT'] = st.session_state.get(f"t2_edit_aot_{ts2_id}", "無資料")
                                        df_map.loc[idx_update, 'FT'] = st.session_state.get(f"t2_edit_ft_{ts2_id}", "無資料")

                                        df_upload = df_map.copy()
                                        df_upload.rename(columns={"TS2#": "NO."}, inplace=True)
                                        df_upload.to_excel("TS2_mapping.xlsx", index=False)
                                        
                                        output = io.BytesIO()
                                        with pd.ExcelWriter(output, engine='openpyxl') as writer: df_upload.to_excel(writer, index=False)
                                        repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                        contents = repo.get_contents("TS2_mapping.xlsx")
                                        repo.update_file(contents.path, f"Update TS2#{ts2_id} via Streamlit", output.getvalue(), contents.sha)
                                        st.cache_data.clear()
                                        
                                        st.success("✅ 成功同步至 GitHub！畫面即將重新載入...")
                                        time.sleep(1.5)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ 上傳失敗: {e}")
                    else:
                        st.caption(f"🔍 站點狀態 👉 JTAG: `{row.get('JTAG', '無資料')}` | AOT: `{row.get('AOT', '無資料')}` | FT: `{row.get('FT', '無資料')}`")
        else:
            st.error(f"⚠️ 找不到資料，請確認輸入是否有誤。")


# ==========================================
# 分頁 3: TS2 STATUS
# ==========================================
with tab_status:
    # ★ 新增：將標題與上傳/下載 TS2_note 檔案區塊並排 ★
    col_title_t3, col_upload_t3 = st.columns([0.6, 0.4])
    with col_title_t3:
        st.header("📊 TS2 STATUS")
    with col_upload_t3:
        with st.expander("📤 上傳 / 下載 TS2_note", expanded=False):
            try:
                with open("TS2_note.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前 Note 檔", data=f, file_name="TS2_note.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError:
                pass
                
            st.divider()
            uploaded_note = st.file_uploader("選擇 Note 檔案 (.xlsx)", type=["xlsx"], key="upload_note")
            if uploaded_note:
                try:
                    df_note_test = pd.read_excel(uploaded_note, dtype=str)
                    
                    if 'NO.' not in df_note_test.columns:
                        st.error("❌ 嚴重錯誤：找不到 'NO.' 欄位，無法解析此檔案。")
                    else:
                        missing_note_cols = [c for c in ['NOTE', 'NOTICE'] if c not in df_note_test.columns]
                        st.success(f"✅ 驗證通過！共讀取到 {len(df_note_test)} 筆備註資料。")
                        if missing_note_cols: st.warning(f"⚠️ 警告：檔案缺少以下欄位 ({', '.join(missing_note_cols)})，系統將在寫入時自動補齊。")
                            
                        if st.button("🚀 確認上傳並覆蓋至 GitHub", key="btn_upload_note", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 上傳中..."):
                                    uploaded_note.seek(0)
                                    excel_bytes = uploaded_note.read()
                                    with open("TS2_note.xlsx", "wb") as f: f.write(excel_bytes)
                                    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                    contents = repo.get_contents("TS2_note.xlsx")
                                    repo.update_file(contents.path, "Update TS2_note.xlsx via Streamlit Upload", excel_bytes, contents.sha)
                                    st.cache_data.clear()
                                    st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                    time.sleep(1.5)
                                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")

    def set_ts2_status_search(num_str):
        st.session_state["status_active_ts2"] = num_str

    def get_t3_state(row):
        ft_val = str(row.get('FT', '無資料')).strip().upper()
        if ft_val == "PASS": return "pass"
        elif ft_val == "FAIL": return "fail"
        else: return "empty"

    ts2_status_states = {}
    for idx, row in df_map.iterrows():
        ts2_id = str(row['TS2#']).strip()
        state = get_t3_state(row)
        if ts2_id not in ts2_status_states:
            ts2_status_states[ts2_id] = state
        elif state == "pass" or (state == "fail" and ts2_status_states[ts2_id] == "empty"):
            ts2_status_states[ts2_id] = state

    ts2_notice_states = {}
    for idx, row in df_note.iterrows():
        ts2_id = str(row['NO.']).strip()
        notice_val = str(row.get('NOTICE', '0')).strip()
        if notice_val.endswith('.0'): notice_val = notice_val[:-2]
        ts2_notice_states[ts2_id] = notice_val

    dynamic_custom_css_t3 = ""
    t3_pass_cnt, t3_fail_cnt, t3_empty_cnt, t3_notice_cnt = 0, 0, 0, 0
    active_status_ts2 = st.session_state.get("status_active_ts2", "")

    for num in range(1, 100):
        state = ts2_status_states.get(str(num), "empty")
        is_selected = (str(num) == active_status_ts2)
        has_notice = (ts2_notice_states.get(str(num)) == '1')
        
        if has_notice: t3_notice_cnt += 1
        elif state == "pass": t3_pass_cnt += 1
        elif state == "fail": t3_fail_cnt += 1
        else: t3_empty_cnt += 1
        
        if not is_selected:
            if has_notice:
                dynamic_custom_css_t3 += f"""
                div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button {{ background-color: #dc3545 !important; border-color: #dc3545 !important; color: #ffffff !important; }}
                div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button:hover {{ background-color: #c82333 !important; border-color: #bd2130 !important; }}
                """
            elif state == "fail":
                dynamic_custom_css_t3 += f"""
                div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }}
                div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({num}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}
                """
            
    if dynamic_custom_css_t3: st.markdown(f"<style>{dynamic_custom_css_t3}</style>", unsafe_allow_html=True)

    panel_title_t3 = f"🎛️ TS2 STATUS 快速面板 (綠色: PASS({t3_pass_cnt}) / 黃色: FAIL({t3_fail_cnt}) / 紅色: NOTICE({t3_notice_cnt}) / 灰色: 無資料({t3_empty_cnt}) / 藍色: 選取)"
    
    with st.expander(panel_title_t3, expanded=True):
        st.markdown('<div class="t3-panel" style="display:none;"></div>', unsafe_allow_html=True)
        cols = st.columns(99)
        for num in range(1, 100):
            state = ts2_status_states.get(str(num), "empty")
            is_selected = (str(num) == active_status_ts2)
            
            if is_selected: btn_type = "tertiary"
            elif state == "pass": btn_type = "primary"
            else: btn_type = "secondary"
                
            cols[num-1].button(str(num), key=f"btn_t3_{num}", on_click=set_ts2_status_search, args=(str(num),), type=btn_type, use_container_width=True)

    if active_status_ts2:
        match_df = df_map[df_map['TS2#'] == active_status_ts2]
        if not match_df.empty:
            for idx, row in match_df.iterrows():
                ts2_id = row['TS2#']
                
                note_row = df_note[df_note['NO.'] == str(ts2_id)]
                val_note = note_row['NOTE'].values[0] if not note_row.empty else "無資料"
                val_note_str = "" if pd.isna(val_note) or val_note == "無資料" else str(val_note)
                val_notice = str(note_row['NOTICE'].values[0]).strip() if not note_row.empty else '0'
                if val_notice.endswith('.0'): val_notice = val_notice[:-2]
                
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.7, 0.3], vertical_alignment="center")
                    with c_title: st.markdown(f"### 🔹 系統標號：TS2#{ts2_id}")
                    with c_toggle: is_editing = st.toggle("✏️ 進入編輯模式", key=f"t3_toggle_{ts2_id}")
                    
                    if val_notice == '1' and not is_editing:
                        st.error("🚨 **此機台已設定特別標註 (NOTICE)**")

                    status_ext_opts = ["無資料", "ongoing", "hold", "NV debug", "OE debug", "Testing", "Other"]
                    def get_ext_status_idx(val):
                        if pd.isna(val) or str(val).strip() == "無資料": return 0
                        v = str(val).strip()
                        for i, opt in enumerate(status_ext_opts):
                            if v.lower() == opt.lower(): return i
                        status_ext_opts.append(v)
                        return len(status_ext_opts) - 1

                    if is_editing:
                        st.markdown(
                            f"**CSM BASE**: <code style='font-size: 18px;'>{row.get('CSM BASE', '無資料')}</code><br>"
                            f"**CSM TRAY**: <code style='font-size: 18px;'>{row.get('CSM TRAY', '無資料')}</code><br>"
                            f"**FULL SYS**: <code style='font-size: 18px;'>{row.get('FULL SYS', '無資料')}</code>",
                            unsafe_allow_html=True
                        )
                        st.markdown("#### 🔍 站點狀態")
                        s1, s2, s3 = st.columns(3)
                        new_jtag = s1.selectbox("JTAG", station_opts, index=get_station_idx(row.get('JTAG')), key=f"t3_edit_jtag_{ts2_id}")
                        new_aot = s2.selectbox("AOT", station_opts, index=get_station_idx(row.get('AOT')), key=f"t3_edit_aot_{ts2_id}")
                        new_ft = s3.selectbox("FT", station_opts, index=get_station_idx(row.get('FT')), key=f"t3_edit_ft_{ts2_id}")

                        st.divider()
                        st.markdown("#### 📋 附加資訊")
                        e1, e2 = st.columns(2)
                        new_status = e1.selectbox("STATUS", status_ext_opts, index=get_ext_status_idx(row.get('STATUS')), key=f"t3_edit_status_{ts2_id}")
                        val_owner = row['OWNER'] if pd.notna(row['OWNER']) and row['OWNER'] != "無資料" else ""
                        new_owner = e2.text_input("OWNER", value=val_owner, key=f"t3_edit_owner_{ts2_id}")
                        
                        val_fail_bin = row.get('Failure BIN', '無資料')
                        val_fail_bin = val_fail_bin if pd.notna(val_fail_bin) and val_fail_bin != "無資料" else ""
                        new_fail_bin = st.text_input("Failure BIN", value=val_fail_bin, key=f"t3_edit_fail_bin_{ts2_id}")

                        st.markdown("#### 📝 備註與標註 (儲存於 TS2_note.xlsx)")
                        new_notice = st.checkbox("🚨 設定為特別標註 (紅色按鈕)", value=(val_notice == '1'), key=f"t3_edit_notice_{ts2_id}")
                        new_note = st.text_area("詳細備註內容", value=val_note_str, height=100, label_visibility="collapsed", key=f"t3_edit_note_{ts2_id}")

                        st.write("")
                        if st.button("💾 儲存修改並同步至 GitHub", key=f"t3_save_btn_{ts2_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新並上傳 Mapping 與 Note 檔案..."):
                                    try:
                                        repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                        
                                        idx_update = df_map[df_map['TS2#'] == ts2_id].index
                                        df_map.loc[idx_update, 'JTAG'] = new_jtag
                                        df_map.loc[idx_update, 'AOT'] = new_aot
                                        df_map.loc[idx_update, 'FT'] = new_ft
                                        df_map.loc[idx_update, 'STATUS'] = new_status
                                        df_map.loc[idx_update, 'OWNER'] = new_owner.strip() or "無資料"
                                        df_map.loc[idx_update, 'Failure BIN'] = new_fail_bin.strip() or "無資料"

                                        df_upload_map = df_map.copy()
                                        df_upload_map.rename(columns={"TS2#": "NO."}, inplace=True)
                                        df_upload_map.to_excel("TS2_mapping.xlsx", index=False)
                                        out_map = io.BytesIO()
                                        with pd.ExcelWriter(out_map, engine='openpyxl') as writer: df_upload_map.to_excel(writer, index=False)
                                        contents_map = repo.get_contents("TS2_mapping.xlsx")
                                        repo.update_file(contents_map.path, f"Update TS2#{ts2_id} Mapping via Tab3", out_map.getvalue(), contents_map.sha)

                                        idx_note = df_note[df_note['NO.'] == str(ts2_id)].index
                                        new_notice_str = '1' if new_notice else '0'
                                        
                                        if not idx_note.empty:
                                            df_note.loc[idx_note, 'NOTE'] = new_note.strip() or "無資料"
                                            df_note.loc[idx_note, 'NOTICE'] = new_notice_str
                                        else:
                                            new_row = pd.DataFrame([{"BUILD": "TS2", "NO.": str(ts2_id), "NOTE": new_note.strip() or "無資料", "NOTICE": new_notice_str}])
                                            df_note = pd.concat([df_note, new_row], ignore_index=True)
                                            
                                        df_note.to_excel("TS2_note.xlsx", index=False)
                                        out_note = io.BytesIO()
                                        with pd.ExcelWriter(out_note, engine='openpyxl') as writer: df_note.to_excel(writer, index=False)
                                        contents_note = repo.get_contents("TS2_note.xlsx")
                                        repo.update_file(contents_note.path, f"Update TS2#{ts2_id} Note via Tab3", out_note.getvalue(), contents_note.sha)

                                        st.cache_data.clear()
                                        st.success("✅ 成功同步 Mapping 與 Note 資料至 GitHub！畫面即將重新載入...")
                                        time.sleep(1.5)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ 上傳失敗: {e}")
                    else:
                        st.markdown(
                            f"**CSM BASE**: <code style='font-size: 18px;'>{row.get('CSM BASE', '無資料')}</code><br>"
                            f"**CSM TRAY**: <code style='font-size: 18px;'>{row.get('CSM TRAY', '無資料')}</code><br>"
                            f"**FULL SYS**: <code style='font-size: 18px;'>{row.get('FULL SYS', '無資料')}</code>",
                            unsafe_allow_html=True
                        )
                        st.markdown(f"🔍 **JTAG**: `{row.get('JTAG', '無資料')}` ｜ **AOT**: `{row.get('AOT', '無資料')}` ｜ **FT**: `{row.get('FT', '無資料')}`")
                        st.divider()
                        st.info(
                            f"**STATUS**: `{row.get('STATUS', '無資料')}`  \n"
                            f"**OWNER**: `{row.get('OWNER', '無資料')}`  \n"
                            f"**Failure BIN**: `{row.get('Failure BIN', '無資料')}`"
                        )
                        
                        st.markdown(f"**📝 詳細備註** (來自 TS2_note.xlsx):  \n> {val_note_str if val_note_str else '無資料'}")
        else:
            st.error(f"⚠️ 找不到該筆資料。")