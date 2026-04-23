import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
import requests
import base64
import hashlib 

# 設定台灣時區 (UTC+8)
tz_tw = timezone(timedelta(hours=8))

# --- 📌 系統記憶體初始化 ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "emp_id" not in st.session_state: st.session_state.emp_id = ""
if "user_name" not in st.session_state: st.session_state.user_name = ""
if "role" not in st.session_state: st.session_state.role = ""
if "must_change_pw" not in st.session_state: st.session_state.must_change_pw = False
if "last_active" not in st.session_state: st.session_state.last_active = datetime.now(tz_tw)
if "form_key" not in st.session_state: st.session_state.form_key = 0

# 📌 訊息提示記憶體 (智慧不洗白防呆)
if "msg_maint" not in st.session_state: st.session_state.msg_maint = ""
if "msg_demo_nt" not in st.session_state: st.session_state.msg_demo_nt = ""
if "msg_demo_v" not in st.session_state: st.session_state.msg_demo_v = ""
if "msg_mach_log" not in st.session_state: st.session_state.msg_mach_log = ""
if "msg_admin_add" not in st.session_state: st.session_state.msg_admin_add = ""

if "load_nt_key" not in st.session_state: st.session_state.load_nt_key = 0
if "load_v_key" not in st.session_state: st.session_state.load_v_key = 0

GAS_URL = "https://script.google.com/macros/s/AKfycbxEVcNlZjjFEmkQmH8Ft-P8mVTSQllsfFF0Khf4YE8lmuOvRQBU8lzocmFs04oMm6g5/exec"

def hash_pw(password): return hashlib.sha256(password.encode()).hexdigest()

def pack_params(param_dict):
    lines = [f"{k}：{v.strip()}" for k, v in param_dict.items() if v and str(v).strip() and str(v).strip() != "nan"]
    return "\n".join(lines) if lines else "無"

def unpack_params(param_str):
    if pd.isna(param_str) or str(param_str).strip() in ["", "無", "nan", "None"]: return {}
    res = {}
    for line in str(param_str).split('\n'):
        if '：' in line:
            k, v = line.split('：', 1)
            res[k.strip()] = v.strip()
    return res

def format_params_html(raw_text):
    if str(raw_text).strip() in ["", "無", "nan", "None", "NaN"]: return ""
    lines = [line for line in str(raw_text).split('\n') if '：' in line]
    return "<br>".join(lines)

def render_lam_inputs(stage_name, key_prefix, fk, defaults=None):
    defaults = defaults or {}
    with st.expander(f"📍 {stage_name} 參數"):
        c1, c2 = st.columns(2)
        with c1: t = st.text_input("溫度 (℃)", value=defaults.get("溫度 (℃)", ""), key=f"{key_prefix}_t_{fk}")
        with c2: t_v = st.text_input("抽真空時間 (sec)", value=defaults.get("抽真空時間 (sec)", ""), key=f"{key_prefix}_tv_{fk}")
        c3, c4 = st.columns(2)
        with c3: p = st.text_input("加壓壓力 (kgf/cm²)", value=defaults.get("加壓壓力 (kgf/cm²)", ""), key=f"{key_prefix}_p_{fk}")
        with c4: t_p = st.text_input("加壓時間 (sec)", value=defaults.get("加壓時間 (sec)", ""), key=f"{key_prefix}_tp_{fk}")
        return {"溫度 (℃)": t, "抽真空時間 (sec)": t_v, "加壓壓力 (kgf/cm²)": p, "加壓時間 (sec)": t_p}

def render_lam3_inputs(stage_name, key_prefix, fk, defaults=None):
    defaults = defaults or {}
    with st.expander(f"📍 {stage_name} 參數 (伺服控制)"):
        modes = ["", "Position", "Press", "Fit"]
        old_mode = defaults.get("控制模式", "")
        m_idx = modes.index(old_mode) if old_mode in modes else 0
        mode = st.selectbox("控制模式", modes, index=m_idx, key=f"{key_prefix}_mode_{fk}")
        st.write("---")
        c1, c2 = st.columns(2)
        with c1: t = st.text_input("溫度 (℃)", value=defaults.get("溫度 (℃)", ""), key=f"{key_prefix}_t_{fk}")
        with c2: t_v = st.text_input("抽真空時間 (sec)", value=defaults.get("抽真空時間 (sec)", ""), key=f"{key_prefix}_tv_{fk}")
        c3, c4 = st.columns(2)
        with c3: thick = st.text_input("目前產品厚度 (mm)", value=defaults.get("目前產品厚度 (mm)", ""), key=f"{key_prefix}_thk_{fk}")
        st.markdown("###### 🎯 模式專屬參數 (未填將自動隱藏)")
        c5, c6, c7 = st.columns(3)
        with c5: pos_v = st.text_input("【Position】厚度補償", value=defaults.get("厚度補償 (Position)", ""), key=f"{key_prefix}_pos_{fk}")
        with c6: press_v = st.text_input("【Press】加壓壓力", value=defaults.get("加壓壓力 (Press)", ""), key=f"{key_prefix}_prs_{fk}")
        with c7: fit_v = st.text_input("【Fit】推進量", value=defaults.get("推進量 (Fit)", ""), key=f"{key_prefix}_fit_{fk}")
        st.write("---")
        c8, c9 = st.columns(2)
        with c8: spd = st.text_input("加壓推速度 (mm/sec)", value=defaults.get("加壓推速度 (mm/sec)", ""), key=f"{key_prefix}_spd_{fk}")
        with c9: t_p = st.text_input("加壓時間 (sec)", value=defaults.get("加壓時間 (sec)", ""), key=f"{key_prefix}_tp_{fk}")
        return {"控制模式": mode, "溫度 (℃)": t, "抽真空時間 (sec)": t_v, "目前產品厚度 (mm)": thick, "厚度補償 (Position)": pos_v, "加壓壓力 (Press)": press_v, "推進量 (Fit)": fit_v, "加壓推速度 (mm/sec)": spd, "加壓時間 (sec)": t_p}

@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_maint = gc.open("設備維修知識庫").worksheet("維修紀錄")
    sheet_demo = gc.open("設備維修知識庫").worksheet("實驗參數")
    sheet_mach = gc.open("設備維修知識庫").worksheet("設備機械履歷")
    sheet_users = gc.open("設備維修知識庫").worksheet("使用者帳號")
    return sheet_maint, sheet_demo, sheet_mach, sheet_users

sheet_maint, sheet_demo, sheet_mach, sheet_users = init_connection()

def upload_image(image_file, file_name):
    base64_image = base64.b64encode(image_file.getvalue()).decode('utf-8')
    payload = {"fileName": file_name, "mimeType": image_file.type, "fileData": base64_image}
    return requests.post(GAS_URL, data=payload).text 

@st.cache_data(ttl=60)
def load_data(mode):
    if mode == "maint": data = sheet_maint.get_all_records()
    elif mode == "users": data = sheet_users.get_all_records()
    elif mode == "machine": data = sheet_mach.get_all_records()
    else: data = sheet_demo.get_all_records()
    df = pd.DataFrame(data)
    if not df.empty and mode != "users": df = df.iloc[::-1].reset_index(drop=True)
    return df

st.markdown("""
<style>
.glide-card { background-color: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 5px; border-left: 6px solid #FFA726; }
.glide-title { font-size: 16px; font-weight: 700; color: #333333; margin-bottom: 4px; }
.glide-subtitle { font-size: 13px; color: #555555; margin-bottom: 10px; line-height: 1.4; }
.glide-tag { background-color: #FFF3E0; color: #E65100; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; margin-bottom: 6px;}
.glide-tag-user { background-color: #E3F2FD; color: #1565C0; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; margin-bottom: 6px;}
.glide-solution { font-size: 13px; color: #D84315; background-color: #FBE9E7; padding: 8px; border-radius: 6px; margin-top: 6px; }
.glide-img { width: 100%; border-radius: 8px; margin-top: 10px; border: 1px solid #eee; }
.diff-alert { background-color: #FFEBEE; border-left: 5px solid #D32F2F; padding: 10px; border-radius: 8px; margin-bottom: 10px; color: #C62828; }
.diff-safe { background-color: #F1F8E9; border-left: 5px solid #4CAF50; padding: 10px; border-radius: 8px; margin-bottom: 10px; color: #2E7D32; }
.calc-yellow { background-color: #FFF9C4; color: #333333; padding: 8px 12px; border-radius: 8px; border-left: 5px solid #FBC02D; font-weight: bold; margin-bottom: 10px; }
.calc-green { background-color: #E8F5E9; padding: 12px; border-radius: 8px; border-left: 6px solid #4CAF50; font-size: 18px; font-weight: bold; color: #2E7D32; margin-top: 10px; }
.hmi-title { font-size: 13px; font-weight: bold; color: #1565C0; margin-bottom: 8px; border-bottom: 1px solid #1565C0; padding-bottom: 3px; margin-top: 15px;}
</style>
""", unsafe_allow_html=True)

if st.session_state.logged_in:
    now = datetime.now(tz_tw)
    if now - st.session_state.last_active > timedelta(minutes=15):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.warning("⏱️ 閒置超過 15 分鐘，自動登出。")
        st.rerun()
    else: st.session_state.last_active = now

# ==========================================
# 🔒 登入系統
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>🔐 設備服務與維護系統</h2>", unsafe_allow_html=True)
    with st.form("login_form"):
        emp_id = st.text_input("工號 (EPM_ID)")
        password = st.text_input("密碼", type="password")
        if st.form_submit_button("登入", use_container_width=True):
            users_df = load_data("users")
            user_record = users_df[users_df['EPM_ID'].astype(str) == emp_id]
            if not user_record.empty and hash_pw(password) == str(user_record.iloc[0]['Password_Hash']):
                st.session_state.logged_in = True
                st.session_state.emp_id, st.session_state.user_name, st.session_state.role = emp_id, str(user_record.iloc[0]['Name']), str(user_record.iloc[0]['Role'])
                if str(user_record.iloc[0]['Is_First_Login']).upper() == 'TRUE': st.session_state.must_change_pw = True
                else: 
                    cell = sheet_users.find(str(emp_id), in_column=1)
                    sheet_users.update_acell(f"F{cell.row}", datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M:%S"))
                st.rerun()
            else: st.error("❌ 工號或密碼錯誤！")

elif st.session_state.must_change_pw:
    with st.form("change_pw_form"):
        st.warning("首次登入請修改密碼")
        new_pw = st.text_input("新密碼", type="password")
        confirm_pw = st.text_input("確認新密碼", type="password")
        if st.form_submit_button("確認修改"):
            if new_pw == confirm_pw and len(new_pw) >= 4:
                cell = sheet_users.find(str(st.session_state.emp_id), in_column=1)
                sheet_users.update_acell(f"C{cell.row}", hash_pw(new_pw))
                sheet_users.update_acell(f"E{cell.row}", "FALSE")
                sheet_users.update_acell(f"F{cell.row}", datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M:%S"))
                st.session_state.must_change_pw = False
                st.rerun()
            else: st.error("⚠️ 密碼不一致或太短")

# ==========================================
# 🔓 主系統運行區塊
# ==========================================
else:
    df_maint = load_data("maint")
    df_demo = load_data("demo")
    df_mach = load_data("machine")
    
    all_cust = []
    if not df_maint.empty: all_cust += df_maint['Customer'].astype(str).tolist()
    if not df_demo.empty: all_cust += df_demo['Customer'].astype(str).tolist()
    unique_cust = sorted(list(set([c.strip() for c in all_cust if c.strip() and c.strip() != 'nan'])))

    with st.sidebar:
        st.success(f"👤 {st.session_state.user_name}")
        app_mode = st.radio("系統模式", ["🔧 現場維修系統", "🧪 DEMO 實驗紀錄", "⚙️ 設備機械履歷", "🧮 產品厚度計算機", "👑 管理員後台"] if st.session_state.role == 'Admin' else ["🔧 現場維修系統", "🧪 DEMO 實驗紀錄", "⚙️ 設備機械履歷", "🧮 產品厚度計算機"])
        if st.button("🔄 重新整理", use_container_width=True): st.cache_data.clear(); st.rerun()
        if st.button("🚪 登出", use_container_width=True):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

    st.markdown(f"<h2>{app_mode}</h2>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # 模式 A：現場維修系統
    # ---------------------------------------------------------
    if app_mode == "🔧 現場維修系統":
        tab1, tab2, tab3 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄", "✏️ 修改我的紀錄"])
        with tab1:
            search_keyword = st.text_input("🔍 關鍵字搜尋")
            filtered_m = df_maint.copy()
            if search_keyword:
                mask = filtered_m.apply(lambda row: row.astype(str).str.contains(search_keyword, case=False).any(), axis=1)
                filtered_m = filtered_m[mask]
            
            st.markdown("#### 📦 批次匯出")
            sel_m_ids = [row['Log_ID'] for idx, row in filtered_m.iterrows() if st.session_state.get(f"chk_m_{row['Log_ID']}", False)]
            if sel_m_ids:
                st.download_button(f"📥 匯出勾選 ({len(sel_m_ids)}筆)", data=filtered_m[filtered_m['Log_ID'].isin(sel_m_ids)].to_csv(index=False).encode('utf-8-sig'), file_name="Maint_Export.csv", type="primary")
            
            for idx, row in filtered_m.iterrows():
                c_card, c_action = st.columns([5, 2])
                with c_action:
                    st.checkbox("🔲 勾選此筆", key=f"chk_m_{row['Log_ID']}")
                    st.download_button("📥 單筆匯出", data=pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig'), file_name=f"{row['Log_ID']}.csv", key=f"dl_m_{row['Log_ID']}")
                with c_card:
                    photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                    st.markdown(f'<div class="glide-card"><div class="glide-title">{row["Component"]} <span style="font-size:12px; color:#999;">({row["Log_ID"]})</span></div><div class="glide-tag">📅 {row["Date"]}</div><div class="glide-tag">🏢 {row["Customer"]}</div><div class="glide-tag">⚙️ {row["Machine_Model"]}</div><div class="glide-tag-user">👤 {row.get("Engineer", "")}</div><br><b>狀況：</b>{row["Issue_Desc"]}<br><b>💡 解法：</b>{row["Solution"]}{photo_html}</div>', unsafe_allow_html=True)
                st.markdown("<hr style='border-top: 2px dashed #ccc; margin: 10px 0px 30px 0px;'>", unsafe_allow_html=True)

        with tab2:
            fk = st.session_state.form_key
            with st.form(f"maint_form_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1: m_c = st.selectbox("選擇客戶", [""] + unique_cust)
                with c2: m_cn = st.text_input("自填客戶")
                m_ma = st.selectbox("機型", ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"])
                m_comp = st.selectbox("異常部件", ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"])
                m_is = st.text_area("問題描述")
                m_so = st.text_area("解決方案")
                m_up = st.file_uploader("🖼️ 照片", type=['jpg','png'])
                
                m_msg = st.empty()
                if st.session_state.msg_maint: m_msg.success(st.session_state.msg_maint); st.session_state.msg_maint = ""
                if st.form_submit_button("💾 送出紀錄"):
                    final_c = m_cn if m_cn else m_c
                    if final_c and m_is:
                        log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M")
                        p_url = upload_image(m_up, f"{log_id}.jpg") if m_up else ""
                        sheet_maint.append_row([log_id, datetime.now(tz_tw).strftime("%Y-%m-%d"), st.session_state.user_name, final_c, m_ma, m_comp, m_is, m_so, p_url])
                        st.session_state.msg_maint = f"✅ 成功單號：{log_id}"; st.session_state.form_key += 1
                        st.cache_data.clear(); st.rerun()
                    else: st.error("⚠️ 資料不完整")

        with tab3:
            my_df_m = df_maint if st.session_state.role == 'Admin' else df_maint[df_maint['Engineer'].astype(str).str.strip() == st.session_state.user_name]
            if my_df_m.empty: st.warning("無紀錄")
            else:
                sel_edit_m = st.selectbox("選擇修改", [""] + [f"{r['Log_ID']} - {r['Customer']}" for idx, r in my_df_m.iterrows()])
                if sel_edit_m:
                    ed_id = sel_edit_m.split(" - ")[0]
                    r_dat = df_maint[df_maint['Log_ID'] == ed_id].iloc[0].to_dict()
                    with st.form(f"ed_m_{ed_id}", clear_on_submit=True):
                        e_dt = st.text_input("日期", value=r_dat.get('Date',''))
                        e_cu = st.text_input("客戶", value=r_dat.get('Customer',''))
                        e_ma = st.text_input("機型", value=r_dat.get('Machine_Model',''))
                        e_co = st.text_input("部件", value=r_dat.get('Component',''))
                        e_is = st.text_area("問題", value=r_dat.get('Issue_Desc',''))
                        e_so = st.text_area("解法", value=r_dat.get('Solution',''))
                        e_up = st.file_uploader("🖼️ 換照片", type=['jpg','png'])
                        if st.form_submit_button("覆蓋"):
                            n_url = upload_image(e_up, f"{ed_id}_e.jpg") if e_up else r_dat.get('Photo_URL','')
                            cell = sheet_maint.find(ed_id, in_column=1)
                            sheet_maint.update(values=[[ed_id, e_dt, r_dat.get('Engineer',''), e_cu, e_ma, e_co, e_is, e_so, n_url]], range_name=f"A{cell.row}:I{cell.row}")
                            st.cache_data.clear(); st.success("更新成功！")

    # ---------------------------------------------------------
    # 模式 B：🧪 DEMO 實驗紀錄 
    # ---------------------------------------------------------
    elif app_mode == "🧪 DEMO 實驗紀錄":
        tab1, tab2, tab3, tab4 = st.tabs(["🔍 參數查詢", "➕ 新增 NT+CVP", "➕ 新增 V-160", "✏️ 修改紀錄"])
        
        with tab1:
            search_d = st.text_input("🔍 搜尋實驗紀錄")
            filtered_d = df_demo.copy()
            if search_d:
                mask = filtered_d.apply(lambda row: row.astype(str).str.contains(search_d, case=False).any(), axis=1)
                filtered_d = filtered_d[mask]
            
            st.markdown("#### 📦 批次匯出區")
            sel_d_ids = [row['Log_ID'] for idx, row in filtered_d.iterrows() if st.session_state.get(f"chk_d_{row['Log_ID']}", False)]
            if sel_d_ids:
                st.download_button(f"📥 匯出勾選 ({len(sel_d_ids)}筆)", data=filtered_d[filtered_d['Log_ID'].isin(sel_d_ids)].to_csv(index=False).encode('utf-8-sig'), file_name="DEMO_Export.csv", type="primary")

            for idx, row in filtered_d.iterrows():
                c_card, c_action = st.columns([5, 2])
                with c_action:
                    st.checkbox("🔲 勾選此筆", key=f"chk_d_{row['Log_ID']}")
                    st.download_button("📥 單筆", data=pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig'), file_name=f"{row['Log_ID']}.csv", key=f"dl_d_{row['Log_ID']}")
                with c_card:
                    p_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                    sub_d = {"板材類型": str(row.get('Substrate_Type','')), "基板尺寸": str(row.get('Substrate_Size','')), "膜材種類": str(row.get('Film_Material','')), "膜材型號": str(row.get('Film_Model',''))}
                    h_sub = format_params_html(pack_params(sub_d))
                    h_pre = format_params_html(row.get('Pre_Lam', ''))
                    h_l1 = format_params_html(row.get('Lam_1st', ''))
                    h_l2 = format_params_html(row.get('Lam_2nd', ''))
                    h_l3 = format_params_html(row.get('Lam_3rd', ''))
                    
                    blk = []
                    if "V160" in str(row.get('Equipment','')).upper() or "V-160" in str(row.get('Equipment','')).upper():
                        if h_l1: blk.append(f"<b>🔹 V-160 參數</b><br>{h_l1}")
                    else:
                        if h_pre: blk.append(f"<b>🔹 預貼</b><br>{h_pre}")
                        if h_l1: blk.append(f"<b>🔹 1st</b><br>{h_l1}")
                        if h_l2: blk.append(f"<b>🔹 2nd</b><br>{h_l2}")
                        if h_l3: blk.append(f"<b>🔹 3rd</b><br>{h_l3}")
                    p_blk = f"<div style='background-color:#F9F9F9; padding:8px; border-radius:6px; font-size:12px; color:#555;'>{'<br><br>'.join(blk)}</div>" if blk else ""
                    s_blk = f"<div class='glide-subtitle'><b>基材/膜材</b><br>{h_sub}</div>" if h_sub else ""

                    st.markdown(f'<div class="glide-card"><div class="glide-title">{row.get("Equipment","")} <span style="font-size:12px; color:#999;">({row["Log_ID"]})</span></div><div class="glide-tag">📅 {row["Date"]}</div><div class="glide-tag">🏢 {row["Customer"]}</div><div class="glide-tag-user">👤 {row.get("Operator","")}</div><div class="glide-tag">📊 自評: {row.get("Self_Eval","")}</div>{s_blk}{p_blk}<br><b>備註：</b>{row["Remarks"]}<br><b>反饋：</b>{row["Feedback"]}{p_html}</div>', unsafe_allow_html=True)
                st.markdown("<hr style='border-top: 2px dashed #ccc; margin: 10px 0px 30px 0px;'>", unsafe_allow_html=True)

        with tab2:
            fk = st.session_state.form_key
            cln_opts = [""] + [f"{r['Log_ID']} - {r['Customer']}" for idx, r in df_demo.iterrows() if "V160" not in str(r.get('Equipment','')).upper()]
            cln_sel = st.selectbox("⚡ 帶入歷史 (NT/CVP)", cln_opts, key=f"cln_nt_{fk}")
            c_dat = df_demo[df_demo['Log_ID'] == cln_sel.split(" - ")[0]].iloc[0].to_dict() if cln_sel else {}

            with st.form(f"d_nt_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1: d_c = st.selectbox("客戶", [""] + unique_cust, index=(unique_cust.index(c_dat.get('Customer'))+1 if c_dat.get('Customer') in unique_cust else 0))
                with c2: d_cn = st.text_input("自填客戶", value=c_dat.get('Customer') if not d_c else "")
                d_eq = st.text_input("機台型號", value=c_dat.get('Equipment', ''))
                
                with st.expander("📍 基材與膜材"):
                    c_s1, c_s2 = st.columns(2)
                    st_o = ["", "PCB", "Wafer", "Glass", "其他"]
                    c_st = str(c_dat.get('Substrate_Type', ''))
                    with c_s1: 
                        d_st = st.selectbox("板材類型", st_o, index=(st_o.index(c_st) if c_st in st_o else (4 if c_st else 0)))
                        d_sto = st.text_input("自填板材", value=c_st if d_st=="其他" else "")
                        d_ss = st.text_input("尺寸厚度", value=c_dat.get('Substrate_Size', ''))
                    fm_o = ["", "ABF", "DAF", "NCF", "PI", "其他"]
                    c_fm = str(c_dat.get('Film_Material', ''))
                    with c_s2: 
                        d_fm = st.selectbox("膜材種類", fm_o, index=(fm_o.index(c_fm) if c_fm in fm_o else (5 if c_fm else 0)))
                        d_fmo = st.text_input("自填膜材", value=c_fm if d_fm=="其他" else "")
                        d_fmod = st.text_input("膜材型號", value=c_dat.get('Film_Model', ''))

                pre_def = unpack_params(c_dat.get('Pre_Lam', ''))
                with st.expander("📍 預貼"):
                    cp1, cp2 = st.columns(2)
                    with cp1: p_t = st.text_input("預貼溫度", value=pre_def.get("預貼溫度 (℃)",""))
                    with cp2: p_s = st.text_input("預貼速度", value=pre_def.get("預貼速度 (m/min)",""))
                    cp3, cp4 = st.columns(2)
                    with cp3: p_p = st.text_input("預貼壓力", value=pre_def.get("預貼壓力 (MPa)",""))
                    with cp4: p_m = st.text_input("留邊量", value=pre_def.get("前後留邊量",""))

                l1_d = render_lam_inputs("1st 壓模", "l1", fk, unpack_params(c_dat.get('Lam_1st', '')))
                l2_d = render_lam_inputs("2nd 壓模", "l2", fk, unpack_params(c_dat.get('Lam_2nd', '')))
                l3_d = render_lam3_inputs("3rd 壓模", "l3", fk, unpack_params(c_dat.get('Lam_3rd', '')))
                
                c_q1, c_q2 = st.columns(2)
                with c_q1: d_qt = st.text_input("數量", value=str(c_dat.get('Qty','')))
                e_opts = ["⚪ 尚未", "🟢 佳", "🟡 普通", "🔴 差"]
                with c_q2: d_ev = st.selectbox("自評", e_opts, index=(e_opts.index(str(c_dat.get('Self_Eval',''))) if str(c_dat.get('Self_Eval','')) in e_opts else 0))
                
                d_re = st.text_area("備註", value=str(c_dat.get('Remarks','')))
                d_fb = st.text_area("反饋", value=str(c_dat.get('Feedback','')))
                d_up = st.file_uploader("🖼️ 照片", type=['jpg','png'])
                
                msg_nt = st.empty()
                if st.session_state.msg_demo_nt: msg_nt.success(st.session_state.msg_demo_nt); st.session_state.msg_demo_nt = ""

                if st.form_submit_button("💾 送出 NT/CVP 紀錄"):
                    f_c = d_cn if d_cn else d_c
                    if f_c and d_eq:
                        lid = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        p_url = upload_image(d_up, f"{lid}.jpg") if d_up else ""
                        f_st = d_sto if d_st == "其他" else d_st
                        f_fm = d_fmo if d_fm == "其他" else d_fm
                        pre_p = pack_params({"預貼溫度 (℃)": p_t, "預貼壓力 (MPa)": p_p, "預貼速度 (m/min)": p_s, "前後留邊量": p_m})
                        sheet_demo.append_row([lid, datetime.now(tz_tw).strftime("%Y-%m-%d"), st.session_state.user_name, f_c, d_eq, f_st, d_ss, f_fm, d_fmod, pre_p, pack_params(l1_d), pack_params(l2_d), pack_params(l3_d), d_qt, d_ev, d_re, d_fb, p_url])
                        st.session_state.msg_demo_nt = f"✅ 成功！單號：{lid}"; st.session_state.form_key += 1
                        st.cache_data.clear(); st.rerun()
                    else: st.error("⚠️ 請填寫完整")

        with tab3:
            fk = st.session_state.form_key
            cln_opts_v = [""] + [f"{r['Log_ID']} - {r['Customer']}" for idx, r in df_demo.iterrows() if "V160" in str(r.get('Equipment','')).upper()]
            cln_sel_v = st.selectbox("⚡ 帶入歷史 (V-160)", cln_opts_v, key=f"cln_v_{fk}")
            c_dat_v = df_demo[df_demo['Log_ID'] == cln_sel_v.split(" - ")[0]].iloc[0].to_dict() if cln_sel_v else {}

            with st.form(f"d_v_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1: v_c = st.selectbox("客戶", [""] + unique_cust, index=(unique_cust.index(c_dat_v.get('Customer'))+1 if c_dat_v.get('Customer') in unique_cust else 0))
                with c2: v_cn = st.text_input("自填客戶", value=c_dat_v.get('Customer') if not v_c else "")
                v_eq = st.text_input("機台型號", value="V-160", disabled=True)
                
                with st.expander("📍 基材與膜材"):
                    c_s1, c_s2 = st.columns(2)
                    st_o = ["", "PCB", "Wafer", "Glass", "其他"]
                    c_st = str(c_dat_v.get('Substrate_Type', ''))
                    with c_s1: 
                        v_st = st.selectbox("板材類型", st_o, index=(st_o.index(c_st) if c_st in st_o else (4 if c_st else 0)), key=f"v_st_{fk}")
                        v_sto = st.text_input("自填板材", value=c_st if v_st=="其他" else "", key=f"v_sto_{fk}")
                        v_ss = st.text_input("尺寸厚度", value=c_dat_v.get('Substrate_Size', ''), key=f"v_ss_{fk}")
                    fm_o = ["", "ABF", "DAF", "NCF", "PI", "其他"]
                    c_fm = str(c_dat_v.get('Film_Material', ''))
                    with c_s2: 
                        v_fm = st.selectbox("膜材種類", fm_o, index=(fm_o.index(c_fm) if c_fm in fm_o else (5 if c_fm else 0)), key=f"v_fm_{fk}")
                        v_fmo = st.text_input("自填膜材", value=c_fm if v_fm=="其他" else "", key=f"v_fmo_{fk}")
                        v_fmod = st.text_input("膜材型號", value=c_dat_v.get('Film_Model', ''), key=f"v_fmod_{fk}")

                v_def = unpack_params(c_dat_v.get('Lam_1st', ''))
                with st.expander("📍 V-160 專屬參數"):
                    cv1, cv2 = st.columns(2)
                    v_mds = ["", "上", "下", "上下"]
                    with cv1: vm = st.selectbox("加壓模式", v_mds, index=(v_mds.index(v_def.get("加壓模式","")) if v_def.get("加壓模式","") in v_mds else 0), key=f"vm_{fk}")
                    with cv2: vtv = st.text_input("下真空時間", value=v_def.get("下真空時間 (sec)",""), key=f"vtv_{fk}")
                    cv3, cv4 = st.columns(2)
                    with cv3: vtt = st.text_input("上溫度", value=v_def.get("上溫度 (℃)",""), key=f"vtt_{fk}")
                    with cv4: vtb = st.text_input("下溫度", value=v_def.get("下溫度 (℃)",""), key=f"vtb_{fk}")
                    vdt = st.text_input("上硅膠墊垂落時間", value=v_def.get("上硅膠墊垂落時間 (sec)",""), key=f"vdt_{fk}")
                    cv5, cv6 = st.columns(2)
                    with cv5: vpt = st.text_input("上氣囊加壓壓力", value=v_def.get("上氣囊加壓壓力 (kgf/cm²)",""), key=f"vpt_{fk}")
                    with cv6: vtpt = st.text_input("上氣囊加壓時間", value=v_def.get("上氣囊加壓時間 (sec)",""), key=f"vtpt_{fk}")
                    vdb = st.text_input("下加壓延遲時間", value=v_def.get("下加壓延遲時間 (sec)",""), key=f"vdb_{fk}")
                    vdrb = st.text_input("下硅膠墊垂落時間", value=v_def.get("下硅膠墊垂落時間 (sec)",""), key=f"vdrb_{fk}")
                    cv7, cv8 = st.columns(2)
                    with cv7: vpb = st.text_input("下加壓壓力", value=v_def.get("下加壓壓力 (kgf/cm²)",""), key=f"vpb_{fk}")
                    with cv8: vtpb = st.text_input("下加壓時間", value=v_def.get("下加壓時間 (sec)",""), key=f"vtpb_{fk}")
                
                c_q1, c_q2 = st.columns(2)
                with c_q1: v_qt = st.text_input("數量", value=str(c_dat_v.get('Qty','')), key=f"v_qt_{fk}")
                e_opts = ["⚪ 尚未", "🟢 佳", "🟡 普通", "🔴 差"]
                with c_q2: v_ev = st.selectbox("自評", e_opts, index=(e_opts.index(str(c_dat_v.get('Self_Eval',''))) if str(c_dat_v.get('Self_Eval','')) in e_opts else 0), key=f"v_ev_{fk}")
                
                v_re = st.text_area("備註", value=str(c_dat_v.get('Remarks','')), key=f"v_re_{fk}")
                v_fb = st.text_area("反饋", value=str(c_dat_v.get('Feedback','')), key=f"v_fb_{fk}")
                v_up = st.file_uploader("🖼️ 照片", type=['jpg','png'], key=f"v_up_{fk}")

                msg_v = st.empty()
                if st.session_state.msg_demo_v: msg_v.success(st.session_state.msg_demo_v); st.session_state.msg_demo_v = ""

                if st.form_submit_button("💾 送出 V-160 紀錄"):
                    f_c = v_cn if v_cn else v_c
                    if f_c:
                        lid = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        p_url = upload_image(v_up, f"{lid}.jpg") if v_up else ""
                        f_st = v_sto if v_st == "其他" else v_st
                        f_fm = v_fmo if v_fm == "其他" else v_fm
                        v_dict = {"加壓模式": vm, "下真空時間 (sec)": vtv, "上溫度 (℃)": vtt, "下溫度 (℃)": vtb, "上硅膠墊垂落時間 (sec)": vdt, "上氣囊加壓壓力 (kgf/cm²)": vpt, "上氣囊加壓時間 (sec)": vtpt, "下加壓延遲時間 (sec)": vdb, "下硅膠墊垂落時間 (sec)": vdrb, "下加壓壓力 (kgf/cm²)": vpb, "下加壓時間 (sec)": vtpb}
                        sheet_demo.append_row([lid, datetime.now(tz_tw).strftime("%Y-%m-%d"), st.session_state.user_name, f_c, "V-160", f_st, v_ss, f_fm, v_fmod, "無", pack_params(v_dict), "無", "無", v_qt, v_ev, v_re, v_fb, p_url])
                        st.session_state.msg_demo_v = f"✅ 成功！單號：{lid}"; st.session_state.form_key += 1
                        st.cache_data.clear(); st.rerun()
                    else: st.error("⚠️ 請填寫完整")

        with tab4:
            my_df_d = df_demo if st.session_state.role == 'Admin' else df_demo[df_demo['Operator'].astype(str).str.strip() == st.session_state.user_name]
            if my_df_d.empty: st.warning("無紀錄")
            else:
                sel_edit_d = st.selectbox("選擇修改", [""] + [f"{r['Log_ID']} - {r['Customer']}" for idx, r in my_df_d.iterrows()])
                if sel_edit_d:
                    ed_id = sel_edit_d.split(" - ")[0]
                    r_dat = df_demo[df_demo['Log_ID'] == ed_id].iloc[0].to_dict()
                    with st.form(f"ed_d_{ed_id}", clear_on_submit=True):
                        e_dt = st.text_input("日期", value=r_dat.get('Date',''))
                        e_cu = st.text_input("客戶", value=r_dat.get('Customer',''))
                        e_re = st.text_area("備註", value=r_dat.get('Remarks',''))
                        e_up = st.file_uploader("🖼️ 換照片", type=['jpg','png'])
                        if st.form_submit_button("覆蓋 (僅供修改基本備註)"):
                            n_url = upload_image(e_up, f"{ed_id}_e.jpg") if e_up else r_dat.get('Photo_URL','')
                            cell = sheet_demo.find(ed_id, in_column=1)
                            sheet_demo.update_acell(f"D{cell.row}", e_cu)
                            sheet_demo.update_acell(f"B{cell.row}", e_dt)
                            sheet_demo.update_acell(f"P{cell.row}", e_re)
                            sheet_demo.update_acell(f"R{cell.row}", n_url)
                            st.cache_data.clear(); st.success("更新成功！")

    # ---------------------------------------------------------
    # 模式 C：⚙️ 設備機械履歷 (100% 同步 HMI 畫面排版 - 對齊 63 欄位)
    # ---------------------------------------------------------
    elif app_mode == "⚙️ 設備機械履歷":
        MACHINE_PARAM_GROUPS = {
            "A區": ["A_D120_下限位置極限", "A_D122_真空位置下限極限", "A_D314_壓合恆定速度", "A_D6064_加速", "A_D6065_減速", "A_D6090_待機位置", "A_D6092_下限位置", "A_D452_壓力異常限值", "A_D170_真空大氣開放時間", "A_D176_不抽真空時轉矩下限", "A_D177_抽真空時轉矩下限", "A_D854_Film咬合保持"],
            "B區": ["B_D40_驅動軸間隔移動量", "B_D6156_加速時間", "B_D6157_減速時間", "B_D46_張力初始", "B_D47_張力初始時定數", "B_D48_品種張力時定數", "B_D714_加速時間", "B_D715_減速時間", "B_D716_Film送帶速度", "B_D717_工序時間", "B_D718_傳送部擋板停止速度", "B_D328_入口異常時間", "B_D514、520_自動運行中擋板上升延遲"],
            "D1區": ["D_D740_壓合台壓力異常範圍", "D_D742_上升傳感器延遲時間", "D_D743_真空大氣開放時間", "D_D746_高壓ON 逆向壓力時間", "D_D747_壓合台必要推力", "D_D748_逆壓壓力", "D_D749_加壓電控閥調節", "D_D750_逆壓電控閥調節", "D_D782_加壓異常時間", "D_D790_下降時殘留壓力排放時間", "D_D791_自動運轉下降阻斷閥打開延遲", "D_D736_1st手動上升高壓輔助ON", "D_D870_增壓閥下限壓力", "D_D872_異常時間"],
            "D2區": ["D_D752_壓合台壓力異常範圍", "D_D754_上升傳感器延遲時間", "D_D755_真空大氣開放時間", "D_D758_高壓ON 逆向壓力時間", "D_D759_壓合台必要推力", "D_D760_逆壓壓力", "D_D761_加壓電控閥調節", "D_D762_逆壓電控閥調節", "D_D783_加壓異常時間", "D_D792_下降時殘留壓力排放時間", "D_D793_自動運轉下降阻斷閥打開延遲", "D_D738_2nd手動上升高壓輔助ON", "D_D873_增壓閥下限壓力", "D_D875_異常時間"],
            "E區": ["E_D460_定位1次定位量", "E_D462_壓力1次定位量", "E_D466_Fit模式SUS接觸搜索1次定位量", "E_D464_Fit控制推進時1次定位量"]
        }

        tab_m1, tab_m2 = st.tabs(["🔍 參數客變比對", "➕ 紀錄機台現況 (100% 畫面排版對應)"])
        
        with tab_m1:
            search_sn = st.text_input("🔍 輸入機台序號 (SN) 查詢：", placeholder="例如: CVP-1500-001")
            if search_sn:
                sn_recs = df_mach[df_mach['Equipment_SN'].astype(str).str.contains(search_sn, case=False)]
                if sn_recs.empty: 
                    st.warning("找不到此機台。")
                else:
                    current = sn_recs.iloc[0].to_dict()
                    factory = sn_recs.iloc[-1].to_dict()
                    previous = sn_recs.iloc[1].to_dict() if len(sn_recs) > 1 else factory
                    
                    st.success(f"✅ 找到 {len(sn_recs)} 筆紀錄。最新紀錄: {current['Date']} (由 {current.get('Engineer', '未登錄')} 更新)")
                    
                    baseline_mode = st.radio("🔄 選擇比對基準：", ["與「原廠設定」比對", "與「前次紀錄」比對"], horizontal=True)
                    if baseline_mode == "與「原廠設定」比對":
                        baseline_data, b_label, b_date = factory, "原廠", factory['Date']
                    else:
                        baseline_data, b_label, b_date = previous, "前次", previous['Date']
                        
                    st.caption(f"📝 基準：{b_label}紀錄 ({b_date} | 工程師: {baseline_data.get('Engineer', '未登錄')})")
                    st.markdown("#### 🛠️ 機械參數差異比對")
                    
                    c1, c2 = st.columns(2)
                    for i, (grp_name, keys) in enumerate(MACHINE_PARAM_GROUPS.items()):
                        col = c1 if i % 2 == 0 else c2
                        with col:
                            with st.expander(f"📍 {grp_name}", expanded=True):
                                for k in keys:
                                    parts = k.split('_')
                                    label = f"{parts[1]} {parts[2]}" if len(parts) >= 3 else parts[1]
                                    v_b, v_c = str(baseline_data.get(k, '')), str(current.get(k, ''))
                                    is_diff = v_b != v_c and v_b != ""
                                    cls = "diff-alert" if is_diff else "diff-safe"
                                    txt = f"🚨 已變更 ({b_label}: {v_b})" if is_diff else f"✅ 與{b_label}相同"
                                    st.markdown(f"<div class='{cls}'><small style='color:#555;'>{label}</small><br><b style='font-size:16px;'>{v_c}</b> <span style='float:right; font-size:12px;'>{txt}</span></div>", unsafe_allow_html=True)
                    st.info(f"📝 **最新客變備註：**\n{current.get('Remarks', '無')}")

        with tab_m2:
            fk = st.session_state.form_key
            st.info("💡 填寫介面已 100% 還原 HMI 畫面分佈，包含您提供的 A 到 BK 共 63 個正確欄位。")
            with st.form(f"mach_log_f_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                m_sn = c1.text_input("機台序號 SN (必填)", key=f"m_sn_{fk}")
                m_cu = c2.selectbox("客戶廠區 (必填)", [""] + unique_cust, key=f"m_cu_{fk}")
                st.write("---")
                
                input_vals = {}
                def mk_input(k):
                    parts = k.split('_')
                    label = f"{parts[1]} {parts[2]}" if len(parts) >= 3 else parts[1]
                    input_vals[k] = st.text_input(label, key=f"m_i_{k}_{fk}")

                # 📍 畫面 A
                with st.expander("📍 畫面 A：機械參數 A (伺服與極限)", expanded=True):
                    colA1, colA2, colA3 = st.columns(3)
                    with colA1:
                        st.markdown("<div class='hmi-title'>▌ 上半部 (位置極限)</div>", unsafe_allow_html=True)
                        for k in ["A_D120_下限位置極限", "A_D122_真空位置下限極限", "A_D6090_待機位置", "A_D6092_下限位置"]: mk_input(k)
                    with colA2:
                        st.markdown("<div class='hmi-title'>▌ 中半部 (速度與轉矩)</div>", unsafe_allow_html=True)
                        for k in ["A_D314_壓合恆定速度", "A_D6064_加速", "A_D6065_減速", "A_D176_不抽真空時轉矩下限", "A_D177_抽真空時轉矩下限"]: mk_input(k)
                    with colA3:
                        st.markdown("<div class='hmi-title'>▌ 右半部 (壓力與時間)</div>", unsafe_allow_html=True)
                        for k in ["A_D452_壓力異常限值", "A_D170_真空大氣開放時間", "A_D854_Film咬合保持"]: mk_input(k)

                # 📍 畫面 B 
                with st.expander("📍 畫面 B：機械參數 B (馬達與傳送)", expanded=True):
                    colB1, colB2, colB3 = st.columns(3)
                    with colB1:
                        st.markdown("<div class='hmi-title'>▌ 驅動輥伺服</div>", unsafe_allow_html=True)
                        for k in ["B_D40_驅動軸間隔移動量", "B_D6156_加速時間", "B_D6157_減速時間"]: mk_input(k)
                    with colB2:
                        st.markdown("<div class='hmi-title'>▌ 出帶伺服</div>", unsafe_allow_html=True)
                        for k in ["B_D46_張力初始", "B_D47_張力初始時定數", "B_D48_品種張力時定數"]: mk_input(k)
                    with colB3:
                        st.markdown("<div class='hmi-title'>▌ 入料傳送</div>", unsafe_allow_html=True)
                        for k in ["B_D714_加速時間", "B_D715_減速時間", "B_D716_Film送帶速度", "B_D717_工序時間", "B_D718_傳送部擋板停止速度", "B_D328_入口異常時間", "B_D514、520_自動運行中擋板上升延遲"]: mk_input(k)

                # 📍 畫面 D 
                with st.expander("📍 畫面 D：機械參數 D (壓合台)", expanded=True):
                    colD1, colD2 = st.columns(2)
                    with colD1:
                        st.markdown("<div class='hmi-title'>▌ 1st 壓合台</div>", unsafe_allow_html=True)
                        for k in MACHINE_PARAM_GROUPS["D1區"]: mk_input(k)
                    with colD2:
                        st.markdown("<div class='hmi-title'>▌ 2nd 壓合台</div>", unsafe_allow_html=True)
                        for k in MACHINE_PARAM_GROUPS["D2區"]: mk_input(k)

                # 📍 畫面 E
                with st.expander("📍 畫面 E：機械參數 E (定位)", expanded=True):
                    st.markdown("<div class='hmi-title'>▌ 4軸伺服壓合 重要數據</div>", unsafe_allow_html=True)
                    colE1, colE2 = st.columns(2)
                    keys_E = MACHINE_PARAM_GROUPS["E區"]
                    with colE1:
                        for k in keys_E[:2]: mk_input(k)
                    with colE2:
                        for k in keys_E[2:]: mk_input(k)
                
                st.write("---")
                m_re = st.text_area("修改原因 / 現場客變備註", placeholder="詳細記錄本次修改了哪些參數，以及修改原因。")
                
                msg_box = st.empty()
                if st.session_state.msg_mach_log:
                    msg_box.success(st.session_state.msg_mach_log)
                    st.session_state.msg_mach_log = ""

                if st.form_submit_button("💾 一鍵儲存 63 欄位設備履歷"):
                    if m_sn and m_cu:
                        with st.spinner("資料打包寫入中..."):
                            log_id = datetime.now(tz_tw).strftime("MACH-%y%m%d-%H%M")
                            date_str = datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M")
                            
                            # 📌 完美依照你給的 A1~BK1 順序寫入！
                            row_data = [log_id, date_str, st.session_state.user_name, m_cu, m_sn]
                            for k in MACHINE_PARAM_GROUPS["A區"]: row_data.append(input_vals[k])
                            for k in MACHINE_PARAM_GROUPS["B區"]: row_data.append(input_vals[k])
                            for k in MACHINE_PARAM_GROUPS["D1區"]: row_data.append(input_vals[k])
                            for k in MACHINE_PARAM_GROUPS["D2區"]: row_data.append(input_vals[k])
                            for k in MACHINE_PARAM_GROUPS["E區"]: row_data.append(input_vals[k])
                            row_data.append(m_re)
                            
                            sheet_mach.append_row(row_data)
                            st.session_state.msg_mach_log = f"✅ 設備 {m_sn} 的 63 項資料已全數建檔成功！"
                            st.session_state.form_key += 1
                            st.cache_data.clear(); st.rerun()
                    else: st.error("⚠️ 機台序號與客戶廠區為必填欄位！")

    # ---------------------------------------------------------
    # 模式 D：🧮 產品厚度計算機
    # ---------------------------------------------------------
    elif app_mode == "🧮 產品厚度計算機":
        st.info("💡 請在下方輸入框填寫測量數值，系統將即時為您運算。**黃色背景**為系統自動計算的結果，**綠色框框**即為應輸入至機台 3rd 的目標產品厚度。")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 【1st 前】測量與設定")
            val_1 = st.number_input("1. 板材厚度 (不含線路、銅柱)", value=0.00, step=0.01, format="%.2f")
            val_2 = st.number_input("2. 板材厚度 (含線路、銅柱)", value=0.00, step=0.01, format="%.2f")
            val_3 = val_2 - val_1
            st.markdown(f"<div class='calc-yellow'>3. 線路、銅柱高：{val_3:.2f}</div>", unsafe_allow_html=True)
            st.write("---")
            val_4 = st.number_input("4. COVER 厚度 (僅供紀錄)", value=0.00, step=0.01, format="%.2f")
            val_5 = st.number_input("5. 膜材 厚度", value=0.00, step=0.01, format="%.2f")
            val_6 = st.number_input("6. PET 厚度", value=0.00, step=0.01, format="%.2f")
            val_7 = val_1 + val_5 + val_6
            val_8 = val_2 + val_5 + val_6
            st.markdown(f"<div class='calc-yellow'>7. 壓合前總厚度 (不含)：{val_7:.2f}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='calc-yellow'>8. 壓合前總厚度 (含)：{val_8:.2f}</div>", unsafe_allow_html=True)

        with col2:
            st.markdown("### 【1st 後】測量與計算")
            val_9 = st.number_input("9. 板材厚度 (不含線路、銅柱)", value=0.00, step=0.01, format="%.2f")
            gap_9 = val_7 - val_9 if val_9 > 0 else 0.0
            st.markdown(f"<div class='calc-yellow'>壓合前後差距 (7 減 9)：{gap_9:.2f}</div>", unsafe_allow_html=True)
            st.write("---")
            val_10 = st.number_input("10. 板材厚度 (含線路、銅柱)", value=0.00, step=0.01, format="%.2f")
            gap_10 = val_8 - val_10 if val_10 > 0 else 0.0
            st.markdown(f"<div class='calc-yellow'>壓合前後差距 (8 減 10)：{gap_10:.2f}</div>", unsafe_allow_html=True)
            val_11 = val_5 - val_3 - gap_10 if val_10 > 0 else 0.0
            st.write("---")
            st.markdown(f"""<div class="calc-green">🎯 輸入 3rd 產品厚度 (對應第 10 項)：{val_10:.2f}</div>""", unsafe_allow_html=True)
            st.markdown(f"""<div class="calc-yellow" style="margin-top:15px; border-left: 5px solid #FF8F00;">11. 膜尚可壓縮量：{val_11:.2f}</div>""", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # 模式 E：👑 管理員後台
    # ---------------------------------------------------------
    elif app_mode == "👑 管理員後台":
        st.subheader("系統管理員專區")
        t1, t2 = st.tabs(["👥 帳號總覽", "➕ 新增帳號"])
        with t1: st.dataframe(load_data("users")[['EPM_ID', 'Name', 'Role', 'Last_Login']], use_container_width=True, hide_index=True)
        with t2:
            fk = st.session_state.form_key
            with st.form(f"adm_f_{fk}", clear_on_submit=False):
                ni = st.text_input("工號"); nn = st.text_input("姓名"); nr = st.selectbox("權限", ["User", "Admin"])
                if st.session_state.msg_admin_add: st.success(st.session_state.msg_admin_add); st.session_state.msg_admin_add = ""
                if st.form_submit_button("建立"):
                    if ni and nn:
                        sheet_users.append_row([ni, nn, hash_pw("123"), nr, "TRUE", ""])
                        st.session_state.msg_admin_add = f"✅ 已建立 {nn}"; st.session_state.form_key += 1
                        st.cache_data.clear(); st.rerun()
