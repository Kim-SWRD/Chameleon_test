import streamlit as st
import pandas as pd
import time
from github import Github
from utils import *  # 導入工具函式

st.set_page_config(page_title="卡美問題與 SN 查詢", layout="wide")
st.markdown(get_custom_css(), unsafe_allow_html=True)

# ==========================================
# ⚙️ 側邊欄：全域 Build 切換器
# ==========================================
st.sidebar.title("⚙️ 系統設定")
build_options = ["TS2", "TS3", "TS4"] # 可以在這裡隨意擴充未來可能出現的 Build
current_build = st.sidebar.selectbox("📌 選擇目前專案 Build", build_options, index=0)
st.sidebar.divider()
st.sidebar.info(f"👉 目前選中：**{current_build}**\n\n系統將自動存取 `{current_build}_mapping.xlsx` 與 `{current_build}_note.xlsx`")

# --- 初始化 Session State ---
if "map_col" not in st.session_state: st.session_state["map_col"] = f"{current_build}#"
if "map_search_input" not in st.session_state: st.session_state["map_search_input"] = ""
if "status_active_sys" not in st.session_state: st.session_state["status_active_sys"] = ""
if "w_add_no_input" not in st.session_state: st.session_state["w_add_no_input"] = ""
if "show_export_list" not in st.session_state: st.session_state["show_export_list"] = False
if "show_export_resolved" not in st.session_state: st.session_state["show_export_resolved"] = False

# --- 動態讀取資料 ---
df_rca = load_rca_data()
df_map = load_mapping_data(current_build)
df_note = load_note_data(current_build)
df_ho = load_handover_data()
df_work = load_work_item_data()

file_mapping = f"{current_build}_mapping.xlsx"
file_note = f"{current_build}_note.xlsx"

valid_sys_list = get_valid_sys_list(df_map)
station_opts_rca = ["無資料", "PASS", "FAIL"]

# 全域計算目前 Build 狀態
sys_status_states, sys_notice_states = get_sys_states(df_map, df_note)

# ==========================================
# 📑 建立頂部切換分頁 (原生 st.tabs)
# ==========================================
tab_rca, tab_map, tab_status, tab_work = st.tabs(["🔍 SOP", "🔄 Mapping", "📊 WIP", "📋 追踨"])

# ==========================================
# 分頁 1: 故障排除 (SOP)
# ==========================================
with tab_rca:
    col_title_rca, col_upload_rca, col_add_rca = st.columns([0.4, 0.4, 0.2])
    with col_title_rca:
        st.header("🔍 SOP (共用)")
        
    with col_upload_rca:
        with st.expander("📤 上傳 / 下載 RCA", expanded=False):
            try:
                with open("RCA.xlsx", "rb") as f:
                    st.download_button(label="📥 下載目前 RCA 檔", data=f, file_name="RCA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: 
                st.warning("目前伺服器上尚未產生 RCA.xlsx")
                
            st.divider()
            uploaded_rca = st.file_uploader("上傳新的 RCA 檔案 (.xlsx)", type=["xlsx"], key="upload_rca")
            if uploaded_rca:
                try:
                    df_rca_test = pd.read_excel(uploaded_rca, dtype=str)
                    st.success(f"✅ 驗證通過！共讀取到 {len(df_rca_test)} 筆 SOP 資料。")
                    if st.button("🚀 確認上傳並覆蓋", key="btn_upload_rca", use_container_width=True, type="primary"):
                        if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                            st.error("❌ 尚未設定系統同步憑證或 Repo！")
                        else:
                            with st.spinner("🔄 上傳中..."):
                                uploaded_rca.seek(0)
                                excel_bytes = uploaded_rca.read()
                                with open("RCA.xlsx", "wb") as f: f.write(excel_bytes)
                                repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                try:
                                    contents = repo.get_contents("RCA.xlsx")
                                    repo.update_file(contents.path, "Update RCA.xlsx via Streamlit Upload", excel_bytes, contents.sha)
                                except Exception:
                                    repo.create_file("RCA.xlsx", "Upload RCA.xlsx via Streamlit Upload", excel_bytes)
                                st.cache_data.clear()
                                st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                time.sleep(1.5)
                                st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")

    with col_add_rca:
        if st.button("➕ 新增 SOP", use_container_width=True, type="primary"):
            st.session_state["show_add_rca"] = not st.session_state.get("show_add_rca", False)

    # === 新增 SOP 表單 ===
    if st.session_state.get("show_add_rca", False):
        with st.container(border=True):
            st.subheader("🆕 新增 SOP 紀錄")
            
            rca_c1, rca_c2 = st.columns(2)
            with rca_c1:
                unique_stations_all = [x for x in df_rca["STATION"].unique() if x != "無資料"] if not df_rca.empty else []
                sel_station = st.selectbox("STATION", unique_stations_all + ["自訂 (請在下方輸入)"])
                if sel_station == "自訂 (請在下方輸入)": new_station = st.text_input("✍️ 自訂 STATION", key="new_custom_station")
                else: new_station = sel_station
                
                new_bincode = st.text_input("BIN_CODE")
                
            with rca_c2:
                new_bin = st.text_input("BIN")
                new_subbin = st.text_input("SUB_BIN")
                
            new_cause = st.text_area("可能原因 (Cause)")
            new_sol = st.text_area("解決方案 (Solution)")
            
            rca_c3, rca_c4 = st.columns(2)
            new_log = rca_c3.text_input("Ref Log")
            new_rev = rca_c4.text_input("REV")
            
            st.write("")
            col_rca_sub, col_rca_can = st.columns(2)
            if col_rca_sub.button("💾 儲存並同步", type="primary", use_container_width=True):
                if not new_bincode.strip() and not new_bin.strip(): 
                    st.error("⚠️️ 請至少填寫 BIN_CODE 或 BIN！")
                elif "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                    st.error("❌ 尚未設定系統同步憑證或 Repo！")
                else:
                    with st.spinner("🔄 上傳中..."):
                        new_row = pd.DataFrame([{
                            'STATION': new_station.strip() or "無資料",
                            'BIN_CODE': new_bincode.strip() or "無資料",
                            'BIN': new_bin.strip() or "無資料",
                            'SUB_BIN': new_subbin.strip() or "無資料",
                            'Possible Cause': new_cause.strip() or "無資料",
                            'Solution': new_sol.strip() or "無資料",
                            'Ref Log': new_log.strip() or "無資料",
                            'REV': new_rev.strip() or "無資料"
                        }])
                        df_rca_updated = pd.concat([df_rca, new_row], ignore_index=True) if not df_rca.empty else new_row
                        save_df_to_github(df_rca_updated, "RCA.xlsx", "RCA.xlsx", "Add new SOP via Streamlit")
                        st.session_state["show_add_rca"] = False
                        st.cache_data.clear()
                        st.success("✅ 已成功新增 SOP！畫面即將重新載入...")
                        time.sleep(1.5)
                        st.rerun()
                        
            if col_rca_can.button("❌ 取消", use_container_width=True):
                st.session_state["show_add_rca"] = False
                st.rerun()

    st.divider()

    if df_rca.empty:
        st.warning("⚠️ 找不到 RCA.xlsx 或資料為空")
    else:
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
                selected_val = st.selectbox("🏷 第二步：請選擇 BIN", unique_bins, key="rca_bin")
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
            
            st.markdown("**🏷️ BIN_CODE:**")
            st.code(display_bin_code, language="plaintext")
            st.markdown("**🏷️️ BIN 全文:**")
            st.code(display_bin, language="plaintext")
            if selected_sub_bin != "無資料":
                st.markdown("**🏷️ SUB_BIN 全文:**")
                st.code(selected_sub_bin, language="plaintext")
            
            final_df = filtered_df[filtered_df["SUB_BIN"] == selected_sub_bin]
            st.markdown(f"### 💡 找到 {len(final_df)} 筆解決方案")
            
            for index, row in final_df.iterrows():
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.8, 0.2], vertical_alignment="center")
                    with c_title: st.markdown(f"#### 📝 紀錄索引號：#{index}")
                    with c_toggle: is_rca_editing = st.toggle("✏ 進入編輯", key=f"rca_edit_tog_{index}")
                    
                    cause_text = str(row['Possible Cause']).replace('\\n', '\n')
                    solution_text = str(row['Solution']).replace('\\n', '\n')
                    log_text = str(row['Ref Log'])
                    rev_text = str(row['REV'])
                    
                    if is_rca_editing:
                        st.markdown("##### 🏷️ 編輯分類標籤")
                        c_st, c_bc = st.columns(2)
                        e_station = c_st.text_input("STATION", value=row.get('STATION', ''), key=f"e_rca_st_{index}")
                        e_bincode = c_bc.text_input("BIN_CODE", value=row.get('BIN_CODE', ''), key=f"e_rca_bc_{index}")
                        
                        c_bin, c_sub = st.columns(2)
                        e_bin = c_bin.text_input("BIN", value=row.get('BIN', ''), key=f"e_rca_b_{index}")
                        e_subbin = c_sub.text_input("SUB_BIN", value=row.get('SUB_BIN', ''), key=f"e_rca_sb_{index}")
                        
                        st.markdown("##### 📝 編輯對策內容")
                        e_cause = st.text_area("🚨 可能原因 (Cause)", value=cause_text, key=f"e_rca_c_{index}")
                        e_sol = st.text_area("✅ 解決方案 (Solution)", value=solution_text, key=f"e_rca_s_{index}")
                        
                        c_log, c_rev = st.columns(2)
                        e_log = c_log.text_input("Log", value=log_text if log_text != "無資料" else "", key=f"e_rca_l_{index}")
                        e_rev = c_rev.text_input("REV", value=rev_text if rev_text != "無資料" else "", key=f"e_rca_r_{index}")
                        
                        st.write("")
                        col_save, col_del = st.columns(2)
                        with col_save:
                            if st.button("💾 儲存修改並同步", key=f"rca_save_{index}", type="primary", use_container_width=True):
                                if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                    st.error("❌ 尚未設定系統同步憑證或 Repo！")
                                else:
                                    with st.spinner("🔄 更新中..."):
                                        df_rca.loc[index, 'STATION'] = e_station.strip() or "無資料"
                                        df_rca.loc[index, 'BIN_CODE'] = e_bincode.strip() or "無資料"
                                        df_rca.loc[index, 'BIN'] = e_bin.strip() or "無資料"
                                        df_rca.loc[index, 'SUB_BIN'] = e_subbin.strip() or "無資料"
                                        df_rca.loc[index, 'Possible Cause'] = e_cause.strip() or "無資料"
                                        df_rca.loc[index, 'Solution'] = e_sol.strip() or "無資料"
                                        df_rca.loc[index, 'Ref Log'] = e_log.strip() or "無資料"
                                        df_rca.loc[index, 'REV'] = e_rev.strip() or "無資料"
                                        
                                        save_df_to_github(df_rca, "RCA.xlsx", "RCA.xlsx", f"Update SOP index {index} via Streamlit")
                                        st.cache_data.clear()
                                        st.success("✅ 已儲存！畫面即將重新載入...")
                                        time.sleep(1.5)
                                        st.rerun()
                        with col_del:
                            if st.button("🗑️ 刪除此筆紀錄", key=f"rca_del_{index}", type="secondary", use_container_width=True):
                                if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                    st.error("❌ 尚未設定系統同步憑證或 Repo！")
                                else:
                                    with st.spinner("🔄 刪除中..."):
                                        df_rca_deleted = df_rca.drop(index)
                                        save_df_to_github(df_rca_deleted, "RCA.xlsx", "RCA.xlsx", f"Delete SOP index {index} via Streamlit")
                                        st.cache_data.clear()
                                        st.success("✅ 已刪除！畫面即將重新載入...")
                                        time.sleep(1.5)
                                        st.rerun()
                    else:
                        display_cause = cause_text.replace('\n', '  \n')
                        display_sol = solution_text.replace('\n', '  \n')
                        
                        st.error(f"**🚨 可能原因 (Cause):**  \n{display_cause}")
                        st.success(f"**✅ 解決方案 (Solution):**  \n{display_sol}")
                        
                        meta_info = []
                        if log_text != "無資料": meta_info.append(f"**Log:** {log_text}")
                        if rev_text != "無資料": meta_info.append(f"**REV:** {rev_text}")
                        if meta_info: st.caption(" | ".join(meta_info))

# ==========================================
# 分頁 2: Mapping
# ==========================================
with tab_map:
    col_title, col_upload = st.columns([0.6, 0.4])
    with col_title:
        st.header(f"🔄 Mapping ({current_build})")
    with col_upload:
        with st.expander(f"📤 上傳 / 下載 {file_mapping}", expanded=False):
            try:
                with open(file_mapping, "rb") as f:
                    st.download_button(label="📥 下載目前 Mapping 檔", data=f, file_name=file_mapping, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: 
                st.warning(f"目前伺服器上尚未產生 {file_mapping}")
                
            st.divider()
            uploaded_file = st.file_uploader("上傳新的 Mapping 檔案 (.xlsx)", type=["xlsx"])
            if uploaded_file:
                try:
                    df_test = pd.read_excel(uploaded_file, dtype=str)
                    rename_test = {}
                    for col in df_test.columns:
                        c_up = str(col).upper()
                        if "NO." in c_up or f"{current_build}#" in c_up or "TS2#" in c_up: rename_test[col] = "NO."
                        elif "CSM BASE" in c_up or "CSM_BASE" in c_up: rename_test[col] = "CSM BASE"
                        elif "CSM TRAY" in c_up or "CSM_TRAY" in c_up: rename_test[col] = "CSM TRAY"
                        elif "FULL SYS" in c_up or "FULL_SYS" in c_up: rename_test[col] = "FULL SYS"
                    df_test.rename(columns=rename_test, inplace=True)
                    
                    if 'NO.' not in df_test.columns:
                        st.error("❌ 嚴重錯誤：找不到 'NO.' 欄位，無法解析此檔案。")
                    else:
                        missing_sn = [req_col for req_col in ['CSM BASE', 'CSM TRAY', 'FULL SYS'] if req_col not in df_test.columns]
                        st.success(f"✅ 驗證通過！共讀取到 {len(df_test)} 筆資料。")
                        if missing_sn: st.warning(f"⚠️ 警告：檔案缺少以下欄位 ({', '.join(missing_sn)})，仍可強制上傳。")
                            
                        if st.button("🚀 確認上傳並覆蓋", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定系統同步憑證或 Repo！")
                            else:
                                with st.spinner("🔄 上傳中..."):
                                    uploaded_file.seek(0)
                                    excel_bytes = uploaded_file.read()
                                    with open(file_mapping, "wb") as f: f.write(excel_bytes)
                                    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                    try:
                                        contents = repo.get_contents(file_mapping)
                                        repo.update_file(contents.path, f"Update {file_mapping} via Streamlit", excel_bytes, contents.sha)
                                    except Exception:
                                        repo.create_file(file_mapping, f"Create {file_mapping} via Streamlit", excel_bytes)
                                    st.cache_data.clear()
                                    st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                    time.sleep(1.5)
                                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")

    def set_sys_search(num_str):
        st.session_state["map_col"] = f"{current_build}#"
        st.session_state["map_search_input"] = num_str

    def get_sn_count(row):
        cnt = 0
        for col in ['CSM BASE', 'CSM TRAY', 'FULL SYS']:
            val = row.get(col, "無資料")
            if pd.notna(val) and str(val).strip() not in ["", "無資料", "nan", "NaN"]: cnt += 1
        return cnt

    sys_sn_counts = {}
    if not df_map.empty:
        for idx, row in df_map.iterrows():
            sys_id = str(row['SYS#']).strip()
            cnt = get_sn_count(row)
            if sys_id not in sys_sn_counts or cnt > sys_sn_counts[sys_id]: sys_sn_counts[sys_id] = cnt

    current_search_col = st.session_state.get("map_col", f"{current_build}#")
    current_search_val = st.session_state.get("map_search_input", "").strip()
    active_sys_numbers = set()
    
    if current_search_val and not df_map.empty:
        query_val = current_search_val
        if current_search_col == f"{current_build}#":
            query_val = query_val.upper().replace(f"{current_build}#", "").replace("TS#", "").strip()
            search_col_internal = "SYS#"
        else:
            search_col_internal = current_search_col
            
        temp_df = df_map[df_map[search_col_internal] == query_val]
        active_sys_numbers = set(temp_df["SYS#"].dropna().astype(str).tolist())

    dynamic_yellow_css_t2 = ""
    full_cnt, partial_cnt, empty_cnt = 0, 0, 0
    
    for idx, sys_val in enumerate(valid_sys_list):
        c = sys_sn_counts.get(sys_val, 0)
        is_selected = (sys_val in active_sys_numbers)
        
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

    panel_title_t2 = f"🎛️ {current_build}# 快速點選面板 (綠色: 完整({full_cnt}) / 黃色: 缺件({partial_cnt}) / 灰色: 無資料({empty_cnt}) / 框線放大: 目前選取)"
    
    with st.expander(panel_title_t2, expanded=True):
        st.markdown('<div class="t2-panel" style="display:none;"></div>', unsafe_allow_html=True)
        if len(valid_sys_list) == 0:
            st.info(f"📂 尚無 {current_build} 資料，請先上傳 {file_mapping}")
        else:
            cols = st.columns(len(valid_sys_list))
            for idx, sys_val in enumerate(valid_sys_list):
                c = sys_sn_counts.get(sys_val, 0)
                btn_type = "primary" if c == 3 else "secondary"
                cols[idx].button(str(sys_val), key=f"btn_t2_{sys_val}", on_click=set_sys_search, args=(sys_val,), type=btn_type, use_container_width=True)

    with st.expander(f"🔍 條件反查 (使用 {current_build}# 或 SN 搜尋)", expanded=False):
        st.markdown(f"輸入 **{current_build}# NO.** (例如: 2), 或是輸入 **CSM BASE, CSM TRAY, FULL SYS** 任意一組 SN，即可互相反查。")
        map_cols = [f"{current_build}#", "CSM BASE", "CSM TRAY", "FULL SYS"]
        col1, col2 = st.columns([1, 2])
        with col1: search_col = st.selectbox("📌 選擇查詢條件", map_cols, key="map_col")
        with col2: search_val = st.text_input(f"✍️ 請輸入 {search_col}", key="map_search_input").strip()
            
    if search_val and not df_map.empty:
        if search_col == f"{current_build}#":
            search_col_internal = "SYS#"
            search_val_proc = search_val.upper().replace(f"{current_build}#", "").replace("TS#", "").strip()
        else:
            search_col_internal = search_col
            search_val_proc = search_val
            
        match_df = df_map[df_map[search_col_internal] == search_val_proc]
        
        if not match_df.empty:
            st.success("✅ 找到對應的 SN 關聯資料！")
            for idx, row in match_df.iterrows():
                sys_id = row['SYS#']
                
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.7, 0.3], vertical_alignment="center")
                    with c_title: st.markdown(f"### 🔹 系統標號：{current_build}#{sys_id}")
                    with c_toggle: is_editing = st.toggle("✏️ 進入編輯模式", key=f"t2_toggle_{sys_id}")
                    
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.markdown("**CSM BASE**")
                        val_base = row['CSM BASE'] if pd.notna(row['CSM BASE']) and row['CSM BASE'] != "無資料" else ""
                        if is_editing: st.text_input("CSM BASE", value=val_base, label_visibility="collapsed", key=f"t2_edit_base_{sys_id}")
                        else: st.code(val_base if val_base else "無資料", language="plaintext")
                            
                    with c2:
                        st.markdown("**CSM TRAY**")
                        val_tray = row['CSM TRAY'] if pd.notna(row['CSM TRAY']) and row['CSM TRAY'] != "無資料" else ""
                        if is_editing: st.text_input("CSM TRAY", value=val_tray, label_visibility="collapsed", key=f"t2_edit_tray_{sys_id}")
                        else: st.code(val_tray if val_tray else "無資料", language="plaintext")
                            
                    with c3:
                        st.markdown("**FULL SYS**")
                        val_full = row['FULL SYS'] if pd.notna(row['FULL SYS']) and row['FULL SYS'] != "無資料" else ""
                        if is_editing: st.text_input("FULL SYS", value=val_full, label_visibility="collapsed", key=f"t2_edit_full_{sys_id}")
                        else: st.code(val_full if val_full else "無資料", language="plaintext")
                    
                    if is_editing:
                        st.divider()
                        st.markdown("#### 🔍 站點狀態")
                        s_c1, s_c2, s_c3 = st.columns(3)
                        with s_c1: st.selectbox("JTAG", station_opts_rca, index=get_station_idx(row.get('JTAG')), key=f"t2_edit_jtag_{sys_id}")
                        with s_c2: st.selectbox("AOT", station_opts_rca, index=get_station_idx(row.get('AOT')), key=f"t2_edit_aot_{sys_id}")
                        with s_c3: st.selectbox("FT", station_opts_rca, index=get_station_idx(row.get('FT')), key=f"t2_edit_ft_{sys_id}")
                            
                        st.write("") 
                        if st.button("💾 儲存修改並同步", key=f"t2_save_btn_{sys_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定系統同步憑證或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新..."):
                                    try:
                                        idx_update = df_map[df_map['SYS#'] == sys_id].index
                                        df_map.loc[idx_update, 'CSM BASE'] = st.session_state.get(f"t2_edit_base_{sys_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'CSM TRAY'] = st.session_state.get(f"t2_edit_tray_{sys_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'FULL SYS'] = st.session_state.get(f"t2_edit_full_{sys_id}", "").strip() or "無資料"
                                        df_map.loc[idx_update, 'JTAG'] = st.session_state.get(f"t2_edit_jtag_{sys_id}", "無資料")
                                        df_map.loc[idx_update, 'AOT'] = st.session_state.get(f"t2_edit_aot_{sys_id}", "無資料")
                                        df_map.loc[idx_update, 'FT'] = st.session_state.get(f"t2_edit_ft_{sys_id}", "無資料")

                                        df_upload = df_map.copy()
                                        df_upload.rename(columns={"SYS#": "NO."}, inplace=True)
                                        save_df_to_github(df_upload, file_mapping, file_mapping, f"Update {current_build}#{sys_id} via Streamlit")
                                        st.cache_data.clear()
                                        st.success("✅ 成功同步！畫面即將重新載入...")
                                        time.sleep(1.5)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ 上傳失敗: {e}")
                    else:
                        st.caption(f"🔍 站點狀態 👉 JTAG: `{row.get('JTAG', '無資料')}` | AOT: `{row.get('AOT', '無資料')}` | FT: `{row.get('FT', '無資料')}`")
        else:
            st.error(f"⚠️ 找不到資料，請確認輸入是否有誤。")

# ==========================================
# 分頁 3: WIP (STATUS)
# ==========================================
with tab_status:
    col_title_t3, col_upload_t3 = st.columns([0.6, 0.4])
    with col_title_t3:
        st.header(f"📊 WIP ({current_build})")
    with col_upload_t3:
        with st.expander(f"📤 上傳 / 下載 {file_note}", expanded=False):
            try:
                with open(file_note, "rb") as f:
                    st.download_button(label="📥 下載目前 Note 檔", data=f, file_name=file_note, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            except FileNotFoundError: pass
                
            st.divider()
            uploaded_note = st.file_uploader("上傳新的 Note 檔案 (.xlsx)", type=["xlsx"], key="upload_note")
            if uploaded_note:
                try:
                    df_note_test = pd.read_excel(uploaded_note, dtype=str)
                    if 'NO.' not in df_note_test.columns:
                        st.error("❌ 嚴重錯誤：找不到 'NO.' 欄位，無法解析此檔案。")
                    else:
                        st.success(f"✅ 驗證通過！共讀取到 {len(df_note_test)} 筆備註資料。")
                        if st.button("🚀 確認上傳並覆蓋", key="btn_upload_note", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定系統同步憑證或 Repo！")
                            else:
                                with st.spinner("🔄 上傳中..."):
                                    uploaded_note.seek(0)
                                    excel_bytes = uploaded_note.read()
                                    with open(file_note, "wb") as f: f.write(excel_bytes)
                                    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
                                    try:
                                        contents = repo.get_contents(file_note)
                                        repo.update_file(contents.path, f"Update {file_note} via Streamlit Upload", excel_bytes, contents.sha)
                                    except Exception:
                                        repo.create_file(file_note, f"Create {file_note} via Streamlit Upload", excel_bytes)
                                    st.cache_data.clear()
                                    st.success("✅ 檔案已成功更新！畫面即將重新載入...")
                                    time.sleep(1.5)
                                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 解析檔案失敗：{e}")
                    
        with st.expander("📤 上傳 / 下載 Handover (共用)", expanded=False):
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
                        if st.button("🚀 確認上傳並覆蓋", key="btn_upload_ho", use_container_width=True, type="primary"):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定系統同步憑證或 Repo！")
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

    def set_sys_status_search(num_str):
        st.session_state["status_active_sys"] = num_str

    dynamic_custom_css_t3 = ""
    t3_pass_cnt, t3_fail_cnt, t3_empty_cnt, t3_notice_cnt = 0, 0, 0, 0
    active_status_sys = st.session_state.get("status_active_sys", "")

    for idx, sys_val in enumerate(valid_sys_list):
        state = sys_status_states.get(sys_val, "empty")
        is_selected = (sys_val == active_status_sys)
        has_notice = (sys_notice_states.get(sys_val) == '1')
        
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

    panel_title_t3 = f"🎛️ {current_build} WIP 面板 (綠色: PASS({t3_pass_cnt}) / 黃色: FAIL({t3_fail_cnt}) / 紅色: NOTICE({t3_notice_cnt}) / 灰色: 無資料({t3_empty_cnt}) / 框線放大: 目前選取)"
    
    with st.expander(panel_title_t3, expanded=True):
        st.markdown('<div class="t3-panel" style="display:none;"></div>', unsafe_allow_html=True)
        if len(valid_sys_list) == 0: st.info(f"📂 尚無 {current_build} 資料")
        else:
            cols = st.columns(len(valid_sys_list))
            for idx, sys_val in enumerate(valid_sys_list):
                state = sys_status_states.get(sys_val, "empty")
                btn_type = "primary" if state == "pass" else "secondary"
                cols[idx].button(str(sys_val), key=f"btn_t3_{sys_val}", on_click=set_sys_status_search, args=(sys_val,), type=btn_type, use_container_width=True)

    with st.expander("🛠 批次設定特別標註 (NOTICE)", expanded=False):
        st.markdown("在這裡可以一次選擇多台系統，大量開啟或關閉紅色標註。")
        c_bulk_sys, c_bulk_action = st.columns([0.7, 0.3])
        
        with c_bulk_sys:
            selected_bulk_sys = st.multiselect("📌 請選擇要批次操作的機台 (可複選)", valid_sys_list, key="bulk_notice_sys")
            
        with c_bulk_action:
            bulk_action = st.radio("🚨 標註動作", ["🔴 開啟 NOTICE (設為紅色)", "⚪ 關閉 NOTICE (取消紅色)"], key="bulk_notice_action")
            
        if st.button("💾 批次執行並同步", key="btn_bulk_notice", type="primary", use_container_width=True):
            if not selected_bulk_sys:
                st.warning("⚠ 請至少選擇一台機台！")
            elif "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                st.error("❌ 尚未設定系統同步憑證或 Repo！")
            else:
                with st.spinner("🔄 批次更新並上傳 Note 檔案中..."):
                    try:
                        new_notice_val = '1' if "開啟" in bulk_action else '0'
                        for sys_id in selected_bulk_sys:
                            idx_note = df_note[df_note['NO.'] == str(sys_id)].index
                            if not idx_note.empty:
                                df_note.loc[idx_note, 'NOTICE'] = new_notice_val
                            else:
                                new_row = pd.DataFrame([{
                                    "BUILD": current_build, 
                                    "NO.": str(sys_id), 
                                    "NOTE": "無資料", 
                                    "NOTICE": new_notice_val
                                }])
                                df_note = pd.concat([df_note, new_row], ignore_index=True)
                                
                        commit_msg = f"Bulk update NOTICE for {len(selected_bulk_sys)} systems via Streamlit"
                        save_df_to_github(df_note, file_note, file_note, commit_msg)
                        
                        st.cache_data.clear()
                        st.success(f"✅ 成功更新 {len(selected_bulk_sys)} 台機台的狀態！畫面即將重新載入...")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 批次上傳失敗: {e}")

    if active_status_sys and not df_map.empty:
        match_df = df_map[df_map['SYS#'] == active_status_sys]
        if not match_df.empty:
            for idx, row in match_df.iterrows():
                sys_id = row['SYS#']
                note_row = df_note[df_note['NO.'] == str(sys_id)]
                val_note = note_row['NOTE'].values[0] if not note_row.empty else "無資料"
                val_note_str = "" if pd.isna(val_note) or val_note == "無資料" else str(val_note)
                val_notice = str(note_row['NOTICE'].values[0]).strip() if not note_row.empty else '0'
                if val_notice.endswith('.0'): val_notice = val_notice[:-2]
                
                with st.container(border=True):
                    c_title, c_toggle = st.columns([0.7, 0.3], vertical_alignment="center")
                    with c_title: st.markdown(f"### 🔹 系統標號：{current_build}#{sys_id}")
                    with c_toggle: is_editing = st.toggle("✏ 進入編輯模式", key=f"t3_toggle_{sys_id}")
                    
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
                        new_jtag = s1.selectbox("JTAG", station_opts_rca, index=get_station_idx(row.get('JTAG')), key=f"t3_edit_jtag_{sys_id}")
                        new_aot = s2.selectbox("AOT", station_opts_rca, index=get_station_idx(row.get('AOT')), key=f"t3_edit_aot_{sys_id}")
                        new_ft = s3.selectbox("FT", station_opts_rca, index=get_station_idx(row.get('FT')), key=f"t3_edit_ft_{sys_id}")

                        st.divider()
                        st.markdown("#### 📋 附加資訊")
                        e1, e2 = st.columns(2)
                        new_status = e1.selectbox("STATUS", status_ext_opts, index=get_ext_status_idx(row.get('STATUS')), key=f"t3_edit_status_{sys_id}")
                        val_owner = row['OWNER'] if pd.notna(row['OWNER']) and row['OWNER'] != "無資料" else ""
                        new_owner = e2.text_input("OWNER", value=val_owner, key=f"t3_edit_owner_{sys_id}")
                        
                        val_fail_bin = row.get('Failure BIN', '無資料')
                        val_fail_bin = val_fail_bin if pd.notna(val_fail_bin) and val_fail_bin != "無資料" else ""
                        new_fail_bin = st.text_input("Failure BIN", value=val_fail_bin, key=f"t3_edit_fail_bin_{sys_id}")

                        st.markdown("#### 📝 備註與標註")
                        new_notice = st.checkbox("🚨 設定為特別標註 (紅色按鈕)", value=(val_notice == '1'), key=f"t3_edit_notice_{sys_id}")
                        new_note = st.text_area("詳細備註內容", value=val_note_str, height=100, label_visibility="collapsed", key=f"t3_edit_note_{sys_id}")

                        render_handover_status(sys_id, current_build, df_ho, is_editing=True)

                        st.write("")
                        if st.button("💾 儲存修改並同步", key=f"t3_save_btn_{sys_id}", type="primary", use_container_width=True):
                            if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
                                st.error("❌ 尚未設定系統同步憑證或 Repo！")
                            else:
                                with st.spinner("🔄 正在更新並上傳 Mapping 與 Note 檔案..."):
                                    try:
                                        idx_update = df_map[df_map['SYS#'] == sys_id].index
                                        df_map.loc[idx_update, 'JTAG'] = new_jtag
                                        df_map.loc[idx_update, 'AOT'] = new_aot
                                        df_map.loc[idx_update, 'FT'] = new_ft
                                        df_map.loc[idx_update, 'STATUS'] = new_status
                                        df_map.loc[idx_update, 'OWNER'] = new_owner.strip() or "無資料"
                                        df_map.loc[idx_update, 'Failure BIN'] = new_fail_bin.strip() or "無資料"

                                        df_upload_map = df_map.copy()
                                        df_upload_map.rename(columns={"SYS#": "NO."}, inplace=True)
                                        save_df_to_github(df_upload_map, file_mapping, file_mapping, f"Update {current_build}#{sys_id} Mapping via Tab3")

                                        idx_note = df_note[df_note['NO.'] == str(sys_id)].index
                                        new_notice_str = '1' if new_notice else '0'
                                        
                                        if not idx_note.empty:
                                            df_note.loc[idx_note, 'NOTE'] = new_note.strip() or "無資料"
                                            df_note.loc[idx_note, 'NOTICE'] = new_notice_str
                                        else:
                                            new_row = pd.DataFrame([{"BUILD": current_build, "NO.": str(sys_id), "NOTE": new_note.strip() or "無資料", "NOTICE": new_notice_str}])
                                            df_note = pd.concat([df_note, new_row], ignore_index=True)
                                            
                                        save_df_to_github(df_note, file_note, file_note, f"Update {current_build}#{sys_id} Note via Tab3")
                                        st.cache_data.clear()
                                        st.success("✅ 成功同步 Mapping 與 Note 資料！畫面即將重新載入...")
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
                            
                        render_handover_status(sys_id, current_build, df_ho, is_editing=False)
        else:
            st.error(f"⚠️ 找不到該筆資料。")

# ==========================================
# 分頁 4: 追踨
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
        st.header(f"📋 追踨 (共用)")
        
    with col_w_upload:
        with st.expander("📤 上/下傳 Work_item (共用)", expanded=False):
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
                            st.error("❌ 尚未設定系統同步憑證或 Repo！")
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
                build_opts = build_options + ["自訂 (請在下方輸入)"]
                default_idx = build_opts.index(current_build) if current_build in build_opts else 0
                sel_build = st.selectbox("BUILD", build_opts, index=default_idx)
                if sel_build == "自訂 (請在下方輸入)": w_build = st.text_input("✍️ 請輸入自訂 BUILD", key="w_custom_build")
                else: w_build = sel_build
                    
            with w_c2:
                w_no = st.text_input("NO. (系統編號)", value=st.session_state.get("w_add_no_input", ""))

            with w_c3:
                station_opts = ["MGMT_JTAG", "MGMT_FT", "AOT", "SYSTEM_FT", "OTHER", "自訂 (請在下方輸入)"]
                sel_station = st.selectbox("Station", station_opts, index=2)
                if sel_station in ["OTHER", "自訂 (請在下方輸入)"]: w_station = st.text_input("✍️ 請輸入自訂 Station", key="w_custom_station")
                else: w_station = sel_station
                    
            dynamic_custom_css_t4 = ""
            active_w_add_no = st.session_state.get("w_add_no_input", "")
            for idx, sys_val in enumerate(valid_sys_list):
                state = sys_status_states.get(sys_val, "empty")
                has_notice = (sys_notice_states.get(sys_val) == '1')
                is_selected = (sys_val == active_w_add_no)
                
                if has_notice:
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ background-color: #dc3545 !important; border-color: #dc3545 !important; color: #ffffff !important; }} div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button:hover {{ background-color: #c82333 !important; border-color: #bd2130 !important; }}"""
                elif state == "fail":
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"] {{ background-color: #ffc107 !important; border-color: #ffc107 !important; color: #000000 !important; }} div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button[kind="secondary"]:hover {{ background-color: #e0a800 !important; border-color: #e0a800 !important; }}"""
                    
                if is_selected:
                    dynamic_custom_css_t4 += f"""div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div:nth-child({idx + 1}) button {{ border: 4px solid #0056b3 !important; box-shadow: 0px 0px 8px 3px rgba(0,86,179,0.6) !important; transform: scale(1.15) !important; position: relative !important; z-index: 99 !important; }}"""
            if dynamic_custom_css_t4: st.markdown(f"<style>{dynamic_custom_css_t4}</style>", unsafe_allow_html=True)

            with st.expander("🎛️ 點擊展開 NO. 快速選擇面板 (💡 顯示為左側欄選中專案的機台)", expanded=False):
                st.markdown('<div class="t4-panel" style="display:none;"></div>', unsafe_allow_html=True)
                if len(valid_sys_list) == 0: st.info(f"📂 尚無資料")
                else:
                    btn_cols = st.columns(len(valid_sys_list))
                    for idx, sys_val in enumerate(valid_sys_list):
                        state = sys_status_states.get(sys_val, "empty")
                        btn_type = "primary" if state == "pass" else "secondary"
                        btn_cols[idx].button(str(sys_val), key=f"btn_w_add_{sys_val}", on_click=set_add_no, args=(str(sys_val),), type=btn_type)

            # 🛠 這裡已升級為 text_area，讓你自由換行！
            w_desc = st.text_area("📝 事項描述 (支援多行輸入)", height=100)
            
            w_c4, w_c5 = st.columns(2)
            w_date = w_c4.text_input("起始日期 (YYYY-MM-DD)", value=time.strftime("%Y-%m-%d"))
            w_owner = w_c5.text_input("經手人員")
            
            st.write("")
            col_submit, col_cancel = st.columns(2)
            with col_submit: w_submitted = st.button("💾 儲存並同步", type="primary", use_container_width=True)
            with col_cancel: w_canceled = st.button("❌ 取消", type="secondary", use_container_width=True)
                
            if w_canceled:
                st.session_state["show_add_work"] = False
                st.rerun()

            if w_submitted:
                if not w_no.strip(): st.error("⚠️ 請填寫或選擇 NO. (系統編號)！")
                elif "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets: st.error("❌ 尚未設定系統同步憑證或 Repo！")
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

    # ==========================================
    # 顯示所有 Build 的追蹤項目 (共用)
    # ==========================================
    df_ongoing = df_work[df_work['工作狀態'] != 'Resolved']
    df_resolved = df_work[df_work['工作狀態'] == 'Resolved']
    
    # ---------------------------
    # 🔥 目前追蹤 (Ongoing) 區塊
    # ---------------------------
    col_ongoing_title, col_ongoing_export = st.columns([0.8, 0.2])
    with col_ongoing_title:
        st.subheader(f"🔥 目前追蹤 ({len(df_ongoing)})")
    with col_ongoing_export:
        if st.button("📋 匯出目前清單", use_container_width=True):
            st.session_state["show_export_list"] = not st.session_state.get("show_export_list", False)

    # === 彈出的匯出清單框塊 ===
    if st.session_state.get("show_export_list", False):
        with st.container(border=True):
            if df_ongoing.empty:
                st.warning("目前無任何追蹤項目。")
            else:
                export_lines = []
                for i, (_, r) in enumerate(df_ongoing.iterrows(), 1):
                    # 處理換行，讓匯出文字依然完美縮排對齊
                    desc_export = str(r.get('事項描述', '')).replace('\n', '\n       ')
                    export_lines.append(f"{i}. {r.get('BUILD', '')}#{r.get('NO.', '')} : {r.get('Station', '')} - {desc_export}")
                
                st.markdown("**📝 快速複製區** (💡 請點擊下方文字框右上角的『複製按鈕』)")
                st.code("\n".join(export_lines), language="plaintext")
        st.write("")

    if df_ongoing.empty: st.info("目前無待處理項目")
        
    for idx, row in df_ongoing.iterrows():
        item_id = row.get('事項編號', f"W-XXX")
        build_val = row.get('BUILD', '無')
        no_val = row.get('NO.', '無')
        station_val = row.get('Station', '無資料')
        
        desc_val = str(row.get('事項描述', '無標題'))
        # 標題欄過濾掉換行，避免破壞排版
        desc_title = desc_val.replace('\n', ' ')
        # 內文框保留換行，轉為 HTML 的換行標籤
        desc_html = desc_val.replace('\n', '<br>')
        
        date_val = str(row.get('起始日期', '')).strip()
        date_tag = f" :gray-background[ {date_val} ]" if date_val and date_val != "無資料" else ""
        display_title = f":orange-background[ {build_val}#{no_val} ] :blue-background[ {station_val} ] {desc_title}{date_tag}"
        
        with st.expander(display_title, expanded=False):
            is_w_edit = st.toggle("✏️ 進入編輯模式", key=f"w_toggle_{item_id}")
            # 將這裡的描述顯示區塊升級，支援換行標籤
            st.markdown(f"<div style='font-size: 16px; font-weight: bold; color: #004085; background-color: #cce5ff; padding: 10px; border-radius: 5px; margin-bottom: 15px; border: 1px solid #b8daff;'>📝 事項描述：<br>{desc_html}</div>", unsafe_allow_html=True)
            
            item_build_val = str(row.get('BUILD', '')).strip().upper()
            cross_ref_no = str(row.get('NO.', '')).strip()
            
            if item_build_val:
                item_df_map = load_mapping_data(item_build_val)
                match_df = item_df_map[item_df_map['SYS#'] == cross_ref_no]
                if not match_df.empty:
                    m_row = match_df.iloc[0]
                    st.markdown(f"**CSM BASE**: `{m_row.get('CSM BASE', '無資料')}`  \n**CSM TRAY**: `{m_row.get('CSM TRAY', '無資料')}`  \n**FULL SYS**: `{m_row.get('FULL SYS', '無資料')}`")
                    st.markdown(f"**JTAG**: `{m_row.get('JTAG', '無資料')}` ｜ **AOT**: `{m_row.get('AOT', '無資料')}` ｜ **FT**: `{m_row.get('FT', '無資料')}`")
                    st.divider()
            
            if is_w_edit:
                e_build = st.text_input("BUILD", value=row.get('BUILD', ''), key=f"e_build_{item_id}")
                e_no = st.text_input("NO.", value=row.get('NO.', ''), key=f"e_no_{item_id}")
                e_station = st.text_input("Station", value=row.get('Station', ''), key=f"e_station_{item_id}")
                
                # 🛠 這裡升級為 text_area，讓編輯時也能換行！
                e_desc = st.text_area("事項描述", value=desc_val, height=100, key=f"e_desc_{item_id}") 
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
                with col_delete: st.button("🗑 刪除此事項", key=f"w_del_{item_id}", use_container_width=True, on_click=cb_delete_work_item, args=(item_id,))

    st.write("")
    st.write("")
    st.divider()

    # ---------------------------
    # ✅ 已解決 (Resolved) 區塊
    # ---------------------------
    col_res_title, col_res_export = st.columns([0.8, 0.2])
    with col_res_title:
        st.subheader(f"✅ 已解決 ({len(df_resolved)})")
    with col_res_export:
        if st.button("📋 匯出已解決清單", use_container_width=True):
            st.session_state["show_export_resolved"] = not st.session_state.get("show_export_resolved", False)

    # === 彈出的已解決匯出清單框塊 ===
    if st.session_state.get("show_export_resolved", False):
        with st.container(border=True):
            if df_resolved.empty:
                st.warning("目前無任何已解決項目。")
            else:
                export_res_lines = []
                for i, (_, r) in enumerate(df_resolved.iterrows(), 1):
                    build_no = f"{r.get('BUILD', '')}#{r.get('NO.', '')}"
                    station = r.get('Station', '')
                    
                    # 處理換行，讓匯出文字依然完美縮排對齊
                    desc_export = str(r.get('事項描述', '')).replace('\n', '\n       ')
                    
                    report_text = str(r.get('回報狀況', '無資料')).strip()
                    if not report_text: report_text = "無資料"
                    report_text = report_text.replace('\n', '\n           ')
                    
                    export_res_lines.append(f"{i}. {build_no} : {station} - {desc_export}")
                    export_res_lines.append(f"       --> {report_text}")
                
                st.markdown("**📝 已解決清單複製區** (💡 請點擊下方文字框右上角的『複製按鈕』)")
                st.code("\n".join(export_res_lines), language="plaintext")
        st.write("")

    if df_resolved.empty: st.info("目前無已解決項目")
        
    for idx, row in df_resolved.iterrows():
        item_id = row.get('事項編號', f"W-XXX")
        build_val = row.get('BUILD', '無')
        no_val = row.get('NO.', '無')
        station_val = row.get('Station', '無資料')
        
        desc_val = str(row.get('事項描述', '無標題'))
        desc_title = desc_val.replace('\n', ' ')
        desc_html = desc_val.replace('\n', '<br>')
        
        date_val = str(row.get('起始日期', '')).strip()
        date_tag = f" :gray-background[ {date_val} ]" if date_val and date_val != "無資料" else ""
        display_title = f":orange-background[ {build_val}#{no_val} ] :blue-background[ {station_val} ] {desc_title}{date_tag}"
        
        with st.expander(display_title, expanded=False):
            st.markdown(f"<div style='font-size: 16px; font-weight: bold; color: #004085; background-color: #cce5ff; padding: 10px; border-radius: 5px; margin-bottom: 15px; border: 1px solid #b8daff;'>📝 事項描述：<br>{desc_html}</div>", unsafe_allow_html=True)
            
            item_build_val = str(row.get('BUILD', '')).strip().upper()
            cross_ref_no = str(row.get('NO.', '')).strip()
            
            if item_build_val:
                item_df_map = load_mapping_data(item_build_val)
                match_df = item_df_map[item_df_map['SYS#'] == cross_ref_no]
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