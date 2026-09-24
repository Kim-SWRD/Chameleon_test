import streamlit as st
import pandas as pd
import time
from github import Github
from utils import *  # 導入工具函式

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="centered")
st.markdown(get_custom_css(), unsafe_allow_html=True)

# --- 初始化 Session State ---
if "map_col" not in st.session_state: st.session_state["map_col"] = "TS2#"
if "map_search_input" not in st.session_state: st.session_state["map_search_input"] = ""
if "status_active_ts2" not in st.session_state: st.session_state["status_active_ts2"] = ""
if "w_add_no_input" not in st.session_state: st.session_state["w_add_no_input"] = ""

# --- 讀取資料 ---
try: df_rca = load_rca_data()
except FileNotFoundError: st.error("找不到 RCA.xlsx"); st.stop()
try: df_map = load_mapping_data()
except FileNotFoundError: st.error("找不到 TS2_mapping.xlsx"); st.stop()
try: df_note = load_note_data()
except FileNotFoundError: st.error("找不到 TS2_note.xlsx"); st.stop()
try: df_ho = load_handover_data()
except FileNotFoundError: st.error("找不到 Chameleon handover status.xlsx"); st.stop()
df_work = load_work_item_data()

valid_ts2_list = get_valid_ts2_list(df_map)
station_opts_rca = ["無資料", "PASS", "FAIL"]

# 全域計算 TS2 STATUS
ts2_status_states, ts2_notice_states = get_ts2_states(df_map, df_note)

# ==========================================
# 📑 建立頂部切換分頁 (原生 st.tabs，享受最完美的換行與紅線效果)
# ==========================================
tab_rca, tab_map, tab_status, tab_work = st.tabs(["🔍 故障排除", "🔄 Mapping查詢", "📊 TS2 STATUS", "📋 追踨問題"])

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
            except FileNotFoundError: pass
                
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
                                    st.success("✅ 檔案已成功更新！畫面即重新載入...")
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
    
    for idx, ts2_val in enumerate(valid_ts2_list):
        c = ts2_sn_counts.get(ts2_val, 0)
        is_selected = (ts2_val in active_ts2_numbers)
        
        if c == 3: full_cnt += 1
        elif c in [1, 2]: partial_cnt += 1
        else: empty_cnt += 1
        
        if c in [1, 2]:
            dynamic_yellow_css_t2 += f"""
            div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }}
            div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}
            """
            
        if is_selected:
            dynamic_yellow_css_t2 += f"""
            div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{
                border: 4px solid #0056b3 !important; box-shadow: 0px 0px 8px 3px rgba(0,86,179,0.6) !important; transform: scale(1.15) !important; position: relative !important; z-index: 99 !important;
            }}
            """
    if dynamic_yellow_css_t2: st.markdown(f"<style>{dynamic_yellow_css_t2}</style>", unsafe_allow_html=True)

    panel_title_t2 = f"🎛️ TS2# 快速點選面板 (綠色: 完整({full_cnt}) / 黃色: 缺件({partial_cnt}) / 灰色: 無資料({empty_cnt}) / 框線放大: 目前選取)"
    
    with st.expander(panel_title_t2, expanded=True):
        st.markdown('<div class="t2-panel" style="display:none;"></div>', unsafe_allow_html=True)
        if len(valid_ts2_list) == 0:
            st.info("尚無 TS2 資料")
        else:
            cols = st.columns(len(valid_ts2_list))
            for idx, ts2_val in enumerate(valid_ts2_list):
                c = ts2_sn_counts.get(ts2_val, 0)
                btn_type = "primary" if c == 3 else "secondary"
                cols[idx].button(str(ts2_val), key=f"btn_t2_{ts2_val}", on_click=set_ts2_search, args=(ts2_val,), type=btn_type, use_container_width=True)

    with st.expander("🔍 條件反查 (使用 TS2# 或 SN 搜尋)", expanded=False):
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
                        with s_c1: st.selectbox("JTAG", station_opts_rca, index=get_station_idx(row.get('JTAG')), key=f"t2_edit_jtag_{ts2_id}")
                        with s_c2: st.selectbox("AOT", station_opts_rca, index=get_station_idx(row.get('AOT')), key=f"t2_edit_aot_{ts2_id}")
                        with s_c3: st.selectbox("FT", station_opts_rca, index=get_station_idx(row.get('FT')), key=f"t2_edit_ft_{ts2_id}")
                            
                        st.write("") 
                        if st.button("💾 儲存修改並同步至 GitHub", key=f"t2_save_btn_{ts2_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新..."):
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
                                        save_df_to_github(df_upload, "TS2_mapping.xlsx", "TS2_mapping.xlsx", f"Update TS2#{ts2_id} via Streamlit")
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
    col_title_t3, col_upload_t3 = st.columns([0.6, 0.4])
    with col_title_t3:
        st.header("📊 TS2 STATUS")
    with col_upload_t3:
        with st.expander("📤 上傳 / 下載 TS2_note", expanded=False):
            try:
                with open("TS2_note.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前 Note 檔", data=f, file_name="TS2_note.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: pass
                
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
                    
        with st.expander("📤 上傳 / 下載 Handover", expanded=False):
            try:
                with open("Chameleon handover status.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前 Handover 檔", data=f, file_name="Chameleon handover status.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: pass
                
            st.divider()
            uploaded_ho = st.file_uploader("選擇 Handover 檔案 (.xlsx)", type=["xlsx"], key="upload_ho")
            if uploaded_ho:
                try:
                    df_ho_test = pd.read_excel(uploaded_ho, dtype=str)
                    missing_ho_cols = [c for c in ['System', 'Failure Description'] if c not in df_ho_test.columns]
                    if missing_ho_cols:
                        st.error(f"❌ 嚴重錯誤：找不到 {', '.join(missing_ho_cols)} 欄位。")
                    else:
                        st.success(f"✅ 驗證通過！共讀取到 {len(df_ho_test)} 筆交接資料。")
                        if st.button("🚀 確認上傳並覆蓋至 GitHub", key="btn_upload_ho", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 上傳中..."):
                                    uploaded_ho.seek(0)
                                    excel_bytes = uploaded_ho.read()
                                    with open("Chameleon handover status.xlsx", "wb") as f: f.write(excel_bytes)
                                    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                    try:
                                        contents = repo.get_contents("Chameleon handover status.xlsx")
                                        repo.update_file(contents.path, "Update Chameleon handover status.xlsx via Streamlit", excel_bytes, contents.sha)
                                    except Exception:
                                        repo.create_file("Chameleon handover status.xlsx", "Upload Chameleon handover status.xlsx via Streamlit", excel_bytes)
                                    st.cache_data.clear()
                                    st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                    time.sleep(1.5)
                                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")

    def set_ts2_status_search(num_str):
        st.session_state["status_active_ts2"] = num_str

    dynamic_custom_css_t3 = ""
    t3_pass_cnt, t3_fail_cnt, t3_empty_cnt, t3_notice_cnt = 0, 0, 0, 0
    active_status_ts2 = st.session_state.get("status_active_ts2", "")

    for idx, ts2_val in enumerate(valid_ts2_list):
        state = ts2_status_states.get(ts2_val, "empty")
        is_selected = (ts2_val == active_status_ts2)
        has_notice = (ts2_notice_states.get(ts2_val) == '1')
        
        if has_notice: t3_notice_cnt += 1
        elif state == "pass": t3_pass_cnt += 1
        elif state == "fail": t3_fail_cnt += 1
        else: t3_empty_cnt += 1
        
        if has_notice:
            dynamic_custom_css_t3 += f"""
            div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ background-color: #dc3545 !important; border-color: #dc3545 !important; color: #ffffff !important; }}
            div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button:hover {{ background-color: #c82333 !important; border-color: #bd2130 !important; }}
            """
        elif state == "fail":
            dynamic_custom_css_t3 += f"""
            div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }}
            div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}
            """
            
        if is_selected:
            dynamic_custom_css_t3 += f"""
            div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ border: 4px solid #0056b3 !important; box-shadow: 0px 0px 8px 3px rgba(0,86,179,0.6) !important; transform: scale(1.15) !important; position: relative !important; z-index: 99 !important; }}
            """
    if dynamic_custom_css_t3: st.markdown(f"<style>{dynamic_custom_css_t3}</style>", unsafe_allow_html=True)

    panel_title_t3 = f"🎛️ TS2 STATUS 快速面板 (綠色: PASS({t3_pass_cnt}) / 黃色: FAIL({t3_fail_cnt}) / 紅色: NOTICE({t3_notice_cnt}) / 灰色: 無資料({t3_empty_cnt}) / 框線放大: 目前選取)"
    
    with st.expander(panel_title_t3, expanded=True):
        st.markdown('<div class="t3-panel" style="display:none;"></div>', unsafe_allow_html=True)
        if len(valid_ts2_list) == 0: st.info("尚無 TS2 資料")
        else:
            cols = st.columns(len(valid_ts2_list))
            for idx, ts2_val in enumerate(valid_ts2_list):
                state = ts2_status_states.get(ts2_val, "empty")
                btn_type = "primary" if state == "pass" else "secondary"
                cols[idx].button(str(ts2_val), key=f"btn_t3_{ts2_val}", on_click=set_ts2_status_search, args=(ts2_val,), type=btn_type, use_container_width=True)

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
                        new_jtag = s1.selectbox("JTAG", station_opts_rca, index=get_station_idx(row.get('JTAG')), key=f"t3_edit_jtag_{ts2_id}")
                        new_aot = s2.selectbox("AOT", station_opts_rca, index=get_station_idx(row.get('AOT')), key=f"t3_edit_aot_{ts2_id}")
                        new_ft = s3.selectbox("FT", station_opts_rca, index=get_station_idx(row.get('FT')), key=f"t3_edit_ft_{ts2_id}")

                        st.divider()
                        st.markdown("#### 📋 附加資訊")
                        e1, e2 = st.columns(2)
                        new_status = e1.selectbox("STATUS", status_ext_opts, index=get_ext_status_idx(row.get('STATUS')), key=f"t3_edit_status_{ts2_id}")
                        val_owner = row['OWNER'] if pd.notna(row['OWNER']) and row['OWNER'] != "無資料" else ""
                        new_owner = e2.text_input("OWNER", value=val_owner, key=f"t3_edit_owner_{ts2_id}")
                        
                        val_fail_bin = row.get('Failure BIN', '無資料')
                        val_fail_bin = val_fail_bin if pd.notna(val_fail_bin) and val_fail_bin != "無資料" else ""
                        new_fail_bin = st.text_input("Failure BIN", value=val_fail_bin, key=f"t3_edit_fail_bin_{ts2_id}")

                        st.markdown("#### 📝 備註與標註")
                        new_notice = st.checkbox("🚨 設定為特別標註 (紅色按鈕)", value=(val_notice == '1'), key=f"t3_edit_notice_{ts2_id}")
                        new_note = st.text_area("詳細備註內容", value=val_note_str, height=100, label_visibility="collapsed", key=f"t3_edit_note_{ts2_id}")

                        render_handover_status(ts2_id, df_ho, is_editing=True)

                        st.write("")
                        if st.button("💾 儲存修改並同步至 GitHub", key=f"t3_save_btn_{ts2_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新並上傳 Mapping 與 Note 檔案..."):
                                    try:
                                        idx_update = df_map[df_map['TS2#'] == ts2_id].index
                                        df_map.loc[idx_update, 'JTAG'] = new_jtag
                                        df_map.loc[idx_update, 'AOT'] = new_aot
                                        df_map.loc[idx_update, 'FT'] = new_ft
                                        df_map.loc[idx_update, 'STATUS'] = new_status
                                        df_map.loc[idx_update, 'OWNER'] = new_owner.strip() or "無資料"
                                        df_map.loc[idx_update, 'Failure BIN'] = new_fail_bin.strip() or "無資料"

                                        df_upload_map = df_map.copy()
                                        df_upload_map.rename(columns={"TS2#": "NO."}, inplace=True)
                                        save_df_to_github(df_upload_map, "TS2_mapping.xlsx", "TS2_mapping.xlsx", f"Update TS2#{ts2_id} Mapping via Tab3")

                                        idx_note = df_note[df_note['NO.'] == str(ts2_id)].index
                                        new_notice_str = '1' if new_notice else '0'
                                        
                                        if not idx_note.empty:
                                            df_note.loc[idx_note, 'NOTE'] = new_note.strip() or "無資料"
                                            df_note.loc[idx_note, 'NOTICE'] = new_notice_str
                                        else:
                                            new_row = pd.DataFrame([{"BUILD": "TS2", "NO.": str(ts2_id), "NOTE": new_note.strip() or "無資料", "NOTICE": new_notice_str}])
                                            df_note = pd.concat([df_note, new_row], ignore_index=True)
                                            
                                        save_df_to_github(df_note, "TS2_note.xlsx", "TS2_note.xlsx", f"Update TS2#{ts2_id} Note via Tab3")
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
                        
                        st.markdown("**📝 詳細備註:**")
                        display_note = val_note_str if val_note_str else '無資料'
                        display_note = display_note.replace('\n', '  \n')
                        if val_notice == '1': st.error(f"**{display_note}**")
                        else: st.info(f"**{display_note}**")
                            
                        render_handover_status(ts2_id, df_ho, is_editing=False)
        else:
            st.error(f"⚠️ 找不到該筆資料。")

# ==========================================
# 分頁 4: 追踨問題 (Work Items)
# ==========================================
with tab_work:
    def cb_update_work_status(item_id, new_status):
        df_w = load_work_item_data().copy()
        idx = df_w[df_w['事項編號'] == item_id].index
        if not idx.empty:
            df_w.loc[idx[0], '工作狀態'] = new_status
            save_df_to_github(df_w, "Work_item.xlsx", "Work_item.xlsx", f"Update status {item_id}")
            st.cache_data.clear()

    def cb_delete_work_item(item_id):
        df_w = load_work_item_data().copy()
        df_w = df_w[df_w['事項編號'] != item_id]
        save_df_to_github(df_w, "Work_item.xlsx", "Work_item.xlsx", f"Delete item {item_id}")
        st.cache_data.clear()

    def set_add_no(val):
        st.session_state["w_add_no_input"] = val

    col_w_title, col_w_upload, col_w_add = st.columns([0.4, 0.4, 0.2])
    with col_w_title:
        st.header("📋 追踨問題")
        
    with col_w_upload:
        with st.expander("📤 上/下傳 Work_item", expanded=False):
            try:
                with open("Work_item.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前檔", data=f, file_name="Work_item.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: pass
                
            st.divider()
            uploaded_work = st.file_uploader("上傳 Work_item (.xlsx)", type=["xlsx"], key="upload_work")
            if uploaded_work:
                try:
                    df_work_test = pd.read_excel(uploaded_work, dtype=str)
                    if '事項編號' not in df_work_test.columns: st.warning("⚠️ 警告：找不到 '事項編號' 欄位，系統將自動補齊。")
                    st.success(f"✅ 驗證通過！共讀取到 {len(df_work_test)} 筆資料。")
                    if st.button("🚀 確認上傳並覆蓋", key="btn_upload_work", use_container_width=True, type="primary"):
                        if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                            st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                        else:
                            with st.spinner("🔄 上傳中..."):
                                uploaded_work.seek(0)
                                excel_bytes = uploaded_work.read()
                                with open("Work_item.xlsx", "wb") as f: f.write(excel_bytes)
                                repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                try:
                                    contents = repo.get_contents("Work_item.xlsx")
                                    repo.update_file(contents.path, "Update Work_item.xlsx via Streamlit Upload", excel_bytes, contents.sha)
                                except Exception:
                                    repo.create_file("Work_item.xlsx", "Upload Work_item.xlsx via Streamlit Upload", excel_bytes)
                                st.cache_data.clear()
                                st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                time.sleep(1.5)
                                st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")
                    
    with col_w_add:
        if st.button("➕ 新增事項", use_container_width=True, type="primary"):
            st.session_state["show_add_work"] = not st.session_state.get("show_add_work", False)
            
    # === 新增事項表單 ===
    if st.session_state.get("show_add_work", False):
        with st.container(border=True):
            st.subheader("🆕 新增追蹤事項")
            
            w_c1, w_c2, w_c3 = st.columns(3)
            with w_c1:
                build_opts = ["TS2", "TS1", "自訂 (請在下方輸入)"]
                sel_build = st.selectbox("BUILD", build_opts, index=0)
                if sel_build == "自訂 (請在下方輸入)": w_build = st.text_input("✍️ 請輸入自訂 BUILD", key="w_custom_build")
                else: w_build = sel_build
                    
            with w_c2:
                w_no = st.text_input("NO. (系統編號)", value=st.session_state.get("w_add_no_input", ""))

            with w_c3:
                station_opts = ["MGMT_JTAG", "MGMT_FT", "AOT", "SYSTEM_FT", "OTHER", "自訂 (請在下方輸入)"]
                sel_station = st.selectbox("Station", station_opts, index=2)
                if sel_station in ["OTHER", "自訂 (請在下方輸入)"]: w_station = st.text_input("✍️ 請輸入自訂 Station", key="w_custom_station")
                else: w_station = sel_station
                    
            # --- 針對分頁 4 動態產生跟分頁 3 完全一樣的顏色邏輯 CSS (.t4-panel) ---
            dynamic_custom_css_t4 = ""
            active_w_add_no = st.session_state.get("w_add_no_input", "")
            for idx, ts2_val in enumerate(valid_ts2_list):
                state = ts2_status_states.get(ts2_val, "empty")
                has_notice = (ts2_notice_states.get(ts2_val) == '1')
                is_selected = (ts2_val == active_w_add_no)
                
                if has_notice:
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ background-color: #dc3545 !important; border-color: #dc3545 !important; color: #ffffff !important; }} div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button:hover {{ background-color: #c82333 !important; border-color: #bd2130 !important; }}"""
                elif state == "fail":
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }} div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}"""
                    
                if is_selected:
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ border: 4px solid #0056b3 !important; box-shadow: 0px 0px 8px 3px rgba(0,86,179,0.6) !important; transform: scale(1.15) !important; position: relative !important; z-index: 99 !important; }}"""
            if dynamic_custom_css_t4: st.markdown(f"<style>{dynamic_custom_css_t4}</style>", unsafe_allow_html=True)

            with st.expander("🎛️ 點擊展開 NO. 快速選擇面板", expanded=False):
                st.markdown('<div class="t4-panel" style="display:none;"></div>', unsafe_allow_html=True)
                if len(valid_ts2_list) == 0: st.info("尚無資料")
                else:
                    btn_cols = st.columns(len(valid_ts2_list))
                    for idx, ts2_val in enumerate(valid_ts2_list):
                        state = ts2_status_states.get(ts2_val, "empty")
                        btn_type = "primary" if state == "pass" else "secondary"
                        btn_cols[idx].button(str(ts2_val), key=f"btn_w_add_{ts2_val}", on_click=set_add_no, args=(str(ts2_val),), type=btn_type)

            w_desc = st.text_input("事項描述")
            w_c4, w_c5 = st.columns(2)
            w_date = w_c4.text_input("起始日期 (YYYY-MM-DD)", value=time.strftime("%Y-%m-%d"))
            w_owner = w_c5.text_input("經手人員")
            
            st.write("")
            col_submit, col_cancel = st.columns(2)
            with col_submit: w_submitted = st.button("💾 儲存並同步至 GitHub", type="primary", use_container_width=True)
            with col_cancel: w_canceled = st.button("❌ 取消", type="secondary", use_container_width=True)
                
            if w_canceled:
                st.session_state["show_add_work"] = False
                st.rerun()

            if w_submitted:
                if not w_no.strip(): st.error("⚠️ 請填寫或選擇 NO. (系統編號)！")
                elif "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets: st.error("❌ 尚未設定 GitHub Token 或 Repo！")
                else:
                    with st.spinner("🔄 上傳中..."):
                        new_id = get_next_work_id(df_work)
                        new_row = pd.DataFrame([{
                            '事項編號': new_id, '工作狀態': 'Ongoing', 'BUILD': w_build.strip(),
                            'NO.': w_no.strip(), '事項描述': w_desc.strip(), '回報狀況': "", 
                            '起始日期': w_date.strip(), '經手人員': w_owner.strip(), 'Station': w_station.strip()
                        }])
                        df_work = pd.concat([df_work, new_row], ignore_index=True)
                        save_df_to_github(df_work.copy(), "Work_item.xlsx", "Work_item.xlsx", f"Add new work item {new_id} via Streamlit")
                        st.session_state["show_add_work"] = False
                        st.session_state["w_add_no_input"] = "" 
                        st.cache_data.clear()
                        st.success(f"✅ 已成功新增事項！畫面即將重新載入...")
                        time.sleep(1.5)
                        st.rerun()

    st.divider()

    df_ongoing = df_work[df_work['工作狀態'] != 'Resolved']
    df_resolved = df_work[df_work['工作狀態'] == 'Resolved']
    
    # ---------------------------
    # 🔥 目前追蹤 (Ongoing) 區塊
    # ---------------------------
    st.subheader(f"🔥 目前追蹤 ({len(df_ongoing)})")
    if df_ongoing.empty: st.info("目前無待處理項目")
        
    for idx, row in df_ongoing.iterrows():
        item_id = row.get('事項編號', f"W-XXX")
        build_val = row.get('BUILD', 'TS2')
        no_val = row.get('NO.', '無')
        station_val = row.get('Station', '無資料')
        desc_val = row.get('事項描述', '無標題')
        date_val = str(row.get('起始日期', '')).strip()
        date_tag = f" :gray-background[ {date_val} ]" if date_val and date_val != "無資料" else ""
        display_title = f":orange-background[ {build_val}#{no_val} ] :blue-background[ {station_val} ] {desc_val}{date_tag}"
        
        with st.expander(display_title, expanded=False):
            is_w_edit = st.toggle("✏️ 進入編輯模式", key=f"w_toggle_{item_id}")
            st.markdown(f"<div style='font-size: 16px; font-weight: bold; color: #004085; background-color: #cce5ff; padding: 10px; border-radius: 5px; margin-bottom: 15px; border: 1px solid #b8daff;'>📝 事項描述：{desc_val}</div>", unsafe_allow_html=True)
            cross_ref_no = str(row.get('NO.', '')).strip()
            match_df = df_map[df_map['TS2#'] == cross_ref_no]
            
            if not match_df.empty:
                m_row = match_df.iloc[0]
                st.markdown(f"**CSM BASE**: `{m_row.get('CSM BASE', '無資料')}`  \n**CSM TRAY**: `{m_row.get('CSM TRAY', '無資料')}`  \n**FULL SYS**: `{m_row.get('FULL SYS', '無資料')}`")
                st.markdown(f"**JTAG**: `{m_row.get('JTAG', '無資料')}` ｜ **AOT**: `{m_row.get('AOT', '無資料')}` ｜ **FT**: `{m_row.get('FT', '無資料')}`")
                st.divider()
            
            if is_w_edit:
                e_build = st.text_input("BUILD", value=row.get('BUILD', ''), key=f"e_build_{item_id}")
                e_no = st.text_input("NO.", value=row.get('NO.', ''), key=f"e_no_{item_id}")
                e_station = st.text_input("Station", value=row.get('Station', ''), key=f"e_station_{item_id}")
                e_desc = st.text_input("事項描述", value=row.get('事項描述', ''), key=f"e_desc_{item_id}") 
                e_report = st.text_area("回報狀況", value=row.get('回報狀況', ''), key=f"e_report_{item_id}")
                e_date = st.text_input("起始日期", value=row.get('起始日期', ''), key=f"e_date_{item_id}")
                e_owner = st.text_input("經手人員", value=row.get('經手人員', ''), key=f"e_owner_{item_id}")
                
                st.write("")
                if st.button("💾 儲存修改", key=f"w_save_{item_id}", type="primary", use_container_width=True):
                    with st.spinner("🔄 更新中..."):
                        target_idx = df_work[df_work['事項編號'] == item_id].index[0]
                        df_work.loc[target_idx, 'BUILD'] = e_build
                        df_work.loc[target_idx, 'NO.'] = e_no
                        df_work.loc[target_idx, 'Station'] = e_station
                        df_work.loc[target_idx, '事項描述'] = e_desc
                        df_work.loc[target_idx, '回報狀況'] = e_report
                        df_work.loc[target_idx, '起始日期'] = e_date
                        df_work.loc[target_idx, '經手人員'] = e_owner
                        save_df_to_github(df_work, "Work_item.xlsx", "Work_item.xlsx", f"Update work item {item_id}")
                        st.cache_data.clear()
                        st.success("✅ 已儲存")
                        time.sleep(1)
                        st.rerun()
            else:
                st.markdown(f"🔹 **機台** : {row.get('BUILD', '')}#{row.get('NO.', '')}  \n🔹 **經手人員** : {row.get('經手人員', '')}")
                report_text = str(row.get('回報狀況', '')).replace('\n', '  \n')
                st.markdown(f"🔹 **回報狀況** :  \n{report_text}")
                
                st.write("")
                col_resolve, col_delete = st.columns(2)
                with col_resolve: st.button("✅ 直接標記已解決", key=f"w_quick_resolve_{item_id}", use_container_width=True, on_click=cb_update_work_status, args=(item_id, 'Resolved'))
                with col_delete: st.button("🗑️ 刪除此事項", key=f"w_del_{item_id}", use_container_width=True, on_click=cb_delete_work_item, args=(item_id,))

    st.write("")
    st.write("")
    st.divider()

    # ---------------------------
    # ✅ 已解決 (Resolved) 區塊
    # ---------------------------
    st.subheader(f"✅ 已解決 ({len(df_resolved)})")
    if df_resolved.empty: st.info("目前無已解決項目")
        
    for idx, row in df_resolved.iterrows():
        item_id = row.get('事項編號', f"W-XXX")
        build_val = row.get('BUILD', 'TS2')
        no_val = row.get('NO.', '無')
        station_val = row.get('Station', '無資料')
        desc_val = row.get('事項描述', '無標題')
        date_val = str(row.get('起始日期', '')).strip()
        date_tag = f" :gray-background[ {date_val} ]" if date_val and date_val != "無資料" else ""
        display_title = f":orange-background[ {build_val}#{no_val} ] :blue-background[ {station_val} ] {desc_val}{date_tag}"
        
        with st.expander(display_title, expanded=False):
            st.markdown(f"<div style='font-size: 16px; font-weight: bold; color: #004085; background-color: #cce5ff; padding: 10px; border-radius: 5px; margin-bottom: 15px; border: 1px solid #b8daff;'>📝 事項描述：{desc_val}</div>", unsafe_allow_html=True)
            cross_ref_no = str(row.get('NO.', '')).strip()
            match_df = df_map[df_map['TS2#'] == cross_ref_no]
            
            if not match_df.empty:
                m_row = match_df.iloc[0]
                st.markdown(f"**CSM BASE**: `{m_row.get('CSM BASE', '無資料')}`  \n**CSM TRAY**: `{m_row.get('CSM TRAY', '無資料')}`  \n**FULL SYS**: `{m_row.get('FULL SYS', '無資料')}`")
                st.markdown(f"**JTAG**: `{m_row.get('JTAG', '無資料')}` ｜ **AOT**: `{m_row.get('AOT', '無資料')}` ｜ **FT**: `{m_row.get('FT', '無資料')}`")
                st.divider()
            
            st.markdown(f"🔹 **機台** : {row.get('BUILD', '')}#{row.get('NO.', '')}  \n🔹 **經手人員** : {row.get('經手人員', '')}")
            report_text = str(row.get('回報狀況', '')).replace('\n', '  \n')
            st.markdown(f"🔹 **回報狀況** :  \n{report_text}")
            
            st.write("")
            col_reopen, col_del = st.columns(2)
            with col_reopen: st.button("🔄 恢復追蹤", key=f"w_reopen_{item_id}", type="secondary", use_container_width=True, on_click=cb_update_work_status, args=(item_id, 'Ongoing'))
            with col_del: st.button("🗑️ 刪除此事項", key=f"w_resolved_del_{item_id}", type="primary", use_container_width=True, on_click=cb_delete_work_item, args=(item_id,))