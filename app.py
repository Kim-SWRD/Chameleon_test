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
                    st.error("⚠ 請至少填寫 BIN_CODE 或 BIN！")
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
            st.markdown("**🏷 BIN 全文:**")
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
                            if st.button("🗑 刪除此筆紀錄", key=f"rca_del_{index}", type="secondary", use_container_width=True):
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
        # 💡 修正 1: 防止 session_state 殘留他版 (如 TS2#) 導致 KeyError
        if "#" in current_search_col:
            query_val = query_val.upper().replace(current_search_col, "").replace(f"{current_build}#", "").replace("TS#", "").strip()
            search_col_internal = "SYS#"
        else:
            search_col_internal = current_search_col
            
        # 💡 修正 2: 增加防呆，確保欄位真的存在於當前的 df_map 中
        if search_col_internal in df_map.columns:
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
            div[data-testid="stExpanderDetails"]:has