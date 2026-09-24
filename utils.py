import streamlit as st
import pandas as pd
import io
import time
import re
from github import Github

def get_custom_css():
    return """
    <style>
    /* 🌟 恢復您要求的頂部距離：設定為 60px */
    .block-container { padding-top: 60px !important; }

    /* 共用按鈕顏色 (Primary 綠色, Tertiary 藍色) */
    button[kind="primary"] { background-color: #28a745 !important; border-color: #28a745 !important; color: white !important; }
    button[kind="primary"]:hover { background-color: #218838 !important; border-color: #1e7e34 !important; }
    button[kind="tertiary"] { background-color: #007bff !important; border-color: #007bff !important; color: white !important; }
    button[kind="tertiary"]:hover { background-color: #0056b3 !important; border-color: #0056b3 !important; }

    /* =========================================================
       🔥 手機版防堆疊 & 防文字裁切 終極裝甲
       ========================================================= */
    /* 強制這兩列在任何螢幕寬度下都必須橫排！打破 Streamlit 的手機端預設限制 */
    div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"],
    div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 8px !important;
    }
    
    /* 強制三個欄位平分寬度 */
    div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"] > div[data-testid="column"],
    div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: 33.33% !important;
        flex: 1 1 33.33% !important;
        min-width: 0 !important; /* 確保不會被擠爆 */
    }

    /* 放寬按鈕內邊距與文字換行限制，確保文字即使在小螢幕也能顯示不被裁成 ... */
    div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"] button,
    div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] button {
        padding: 5px 2px !important; /* 縮小左右 padding 騰出文字空間 */
        height: 100% !important;
    }
    
    div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"] button p,
    div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] button p {
        font-size: 14px !important;
        white-space: normal !important; /* 🌟 允許文字自動換行，避免出現 ... */
        word-break: keep-all !important;
        line-height: 1.2 !important;
        text-align: center !important;
    }
    
    @media (min-width: 768px) {
        /* 電腦端螢幕夠大，恢復較大字體與單行顯示 */
        div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"] button p,
        div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] button p {
            font-size: 16px !important;
            white-space: nowrap !important;
        }
        div.element-container:has(.mobile-nav-row1) + div[data-testid="stHorizontalBlock"] button,
        div.element-container:has(.mobile-nav-row2) + div[data-testid="stHorizontalBlock"] button {
            padding: 8px 12px !important;
        }
    }

    /* =========================================================
       網格面板與其他樣式修正 
       ========================================================= */
    div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"],
    div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"],
    div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] {
        display: grid !important; grid-template-columns: repeat(10, 1fr) !important; gap: 4px !important; width: 100% !important; padding-bottom: 3px !important;
    }
    div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] > div[data-testid="column"],
    div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] > div[data-testid="column"],
    div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] > div[data-testid="column"] { 
        width: 100% !important; min-width: 0 !important; padding: 0 !important; 
    }
    div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] button,
    div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] button,
    div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] button {
        width: 200% !important; aspect-ratio: 1 / 1 !important; border-radius: 6px !important; padding: 0 !important; margin: 0 !important; min-height: 0 !important; height: auto !important; display: flex !important; align-items: center !important; justify-content: center !important; transition: all 0.2s ease-in-out !important;
    }
    div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] button > div,
    div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] button > div,
    div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] button > div { 
        display: flex !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; margin: 0 !important; padding: 0 !important; 
    }
    div[data-testid="stExpanderDetails"]:has(.t2-panel) div[data-testid="stHorizontalBlock"] button p,
    div[data-testid="stExpanderDetails"]:has(.t3-panel) div[data-testid="stHorizontalBlock"] button p,
    div[data-testid="stExpanderDetails"]:has(.t4-panel) div[data-testid="stHorizontalBlock"] button p { 
        font-size: 13px !important; font-weight: 700 !important; margin: 0 !important; padding: 0 !important; text-align: center !important; 
    }
    div[data-testid="stCodeBlock"] button { opacity: 1 !important; visibility: visible !important; display: inline-flex !important; }

    /* 🔥 客製化追蹤問題 (Expander) 標題內的色塊顏色 🔥 */
    div[data-testid="stExpander"] summary p span:nth-of-type(1) { background-color: #ffc107 !important; color: #000000 !important; border-radius: 4px !important; padding: 2px 8px !important; font-weight: bold !important; }
    div[data-testid="stExpander"] summary p span:nth-of-type(2) { background-color: #dbeafe !important; color: #1e3a8a !important; border-radius: 4px !important; padding: 2px 8px !important; }
    div[data-testid="stExpander"] summary p span:nth-of-type(3) { background-color: #f1f3f5 !important; color: #6c757d !important; border-radius: 4px !important; padding: 2px 8px !important; font-size: 0.85em !important; margin-left: 6px !important; }
    </style>
    """

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

@st.cache_data(ttl=60)
def load_handover_data():
    df = pd.read_excel("Chameleon handover status.xlsx", dtype=str)
    if 'System' in df.columns and 'Failure Description' in df.columns:
        df_ho = df[['System', 'Failure Description']].dropna(subset=['System'])
        df_ho['System'] = df_ho['System'].astype(str).str.strip()
        df_ho['Failure Description'] = df_ho['Failure Description'].fillna("無資料")
        return df_ho
    return pd.DataFrame(columns=['System', 'Failure Description'])

@st.cache_data(ttl=60)
def load_work_item_data():
    try:
        df = pd.read_excel("Work_item.xlsx", dtype=str)
        if 'NO.' in df.columns:
            df['NO.'] = df['NO.'].astype(str).str.replace(r'\.0$', '', regex=True)
        if '事項編號' not in df.columns:
            df.insert(0, '事項編號', [f"W-{i:03d}" for i in range(1, len(df) + 1)])
            try: df.to_excel("Work_item.xlsx", index=False)
            except PermissionError: pass
        df.fillna("無資料", inplace=True)
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=['事項編號', '工作狀態', 'BUILD', 'NO.', '事項描述', '回報狀況', '起始日期', '經手人員', 'Station'])

def get_next_work_id(df):
    if df.empty or '事項編號' not in df.columns: return "W-001"
    ids = df['事項編號'].astype(str).str.extract(r'W-(\d+)')[0].dropna().astype(int)
    if ids.empty: return "W-001"
    return f"W-{ids.max() + 1:03d}"

def render_handover_status(ts2_id, df_ho, is_editing=False):
    pattern = f"^TS2.*#\\s*{ts2_id}(?:[^0-9]|$)"
    matched_ho = df_ho[df_ho['System'].str.match(pattern, na=False)]
    if is_editing: st.markdown("#### 🔄 日夜交接狀態 (唯讀)")
    else: st.markdown("**🔄 日夜交接狀態 (唯讀):**")
    if matched_ho.empty: st.caption("（無相關交接紀錄）")
    else:
        ho_desc = "\n---\n".join(matched_ho['Failure Description'].astype(str).tolist())
        st.code(ho_desc, language="plaintext")

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

def get_valid_ts2_list(df_map):
    valid_set = set()
    for v in df_map["TS2#"].dropna().astype(str):
        if v.strip() and v.strip() != "無資料" and v.strip().lower() != "nan": valid_set.add(v.strip())
    return sorted(list(valid_set), key=natural_sort_key)

def get_station_idx(val):
    if pd.isna(val) or str(val).strip() == "無資料": return 0
    v = str(val).strip().upper()
    if v == "PASS": return 1
    if v == "FAIL": return 2
    return 0

def save_df_to_github(df, filename, repo_path, commit_message):
    try:
        df.to_excel(filename, index=False)
    except Exception:
        pass 

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    
    repo = Github(st.secrets["GITHUB_TOKEN"]).get_repo(st.secrets["GITHUB_REPO"])
    try:
        contents = repo.get_contents(repo_path)
        repo.update_file(contents.path, commit_message, excel_data, contents.sha)
    except Exception:
        repo.create_file(repo_path, commit_message, excel_data)

def get_ts2_states(df_map, df_note):
    def get_t3_state(row):
        ft_val = str(row.get('FT', '無資料')).strip().upper()
        if ft_val == "PASS": return "pass"
        elif ft_val == "FAIL": return "fail"
        return "empty"

    status_states = {}
    for _, row in df_map.iterrows():
        ts2_id = str(row['TS2#']).strip()
        state = get_t3_state(row)
        if ts2_id not in status_states or state == "pass" or (state == "fail" and status_states[ts2_id] == "empty"):
            status_states[ts2_id] = state

    notice_states = {}
    for _, row in df_note.iterrows():
        ts2_id = str(row['NO.']).strip()
        n_val = str(row.get('NOTICE', '0')).strip()
        if n_val.endswith('.0'): n_val = n_val[:-2]
        notice_states[ts2_id] = n_val

    return status_states, notice_states