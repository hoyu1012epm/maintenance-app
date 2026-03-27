import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
import requests
import base64
import plotly.express as px
import hashlib 

# 設定台灣時區 (UTC+8)
tz_tw = timezone(timedelta(hours=8))

# --- 📌 登入與閒置計時記憶體 ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "emp_id" not in st.session_state:
    st.session_state.emp_id = ""
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "role" not in st.session_state:
    st.session_state.role = ""
if "must_change_pw" not in st.session_state:
    st.session_state.must_change_pw = False
if "last_active" not in st.session_state:
    st.session_state.last_active = datetime.now(tz_tw)
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# 📌 Google Apps Script 專屬接收站網址
GAS_URL = "https://script.google.com/macros/s/AKfycbxEVcNlZjjFEmkQmH8Ft-P8mVTSQllsfFF0Khf4YE8lmuOvRQBU8lzocmFs04oMm6g5/exec"

def hash_pw(password):
    return hashlib.sha256(password.encode()).hexdigest()

def pack_params(param_dict):
    lines = []
    for k, v in param_dict.items():
        if v and str(v).strip() and str(v).strip() != "nan":
            lines.append(f"{k}：{v.strip()}")
    return "\n".join(lines) if lines else "無"

def unpack_params(param_str):
    if pd.isna(param_str) or str(param_str).strip() in ["", "無", "nan", "None"]:
        return {}
    res = {}
    for line in str(param_str).split('\n'):
        if '：' in line:
            k, v = line.split('：', 1)
            res[k.strip()] = v.strip()
    return res

def format_params_html(raw_text):
    if str(raw_text).strip() in ["", "無", "nan", "None", "NaN"]:
        return ""
    lines = str(raw_text).replace('\r', '').split('\n')
    valid_lines = []
    for line in lines:
        if not line.strip(): continue
        if '：' in line and not line.split('：', 1)[1].strip(): continue
        if ':' in line and not line.split(':', 1)[1].strip(): continue
        valid_lines.append(line)
    return "<br>".join(valid_lines) if valid_lines else ""

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
        return {
            "控制模式": mode, "溫度 (℃)": t, "抽真空時間 (sec)": t_v, "目前產品厚度 (mm)": thick, 
            "厚度補償 (Position)": pos_v, "加壓壓力 (Press)": press_v, "推進量 (Fit)": fit_v,
            "加壓推速度 (mm/sec)": spd, "加壓時間 (sec)": t_p
        }

@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_maint = gc.open("設備維修知識庫").worksheet("維修紀錄")
    sheet_demo = gc.open("設備維修知識庫").worksheet("實驗參數")
    sheet_users = gc.open("設備維修知識庫").worksheet("使用者帳號")
    return sheet_maint, sheet_demo, sheet_users

sheet_maint, sheet_demo, sheet_users = init_connection()

def upload_image(image_file, file_name):
    base64_image = base64.b64encode(image_file.getvalue()).decode('utf-8')
    payload = {"fileName": file_name, "mimeType": image_file.type, "fileData": base64_image}
    return requests.post(GAS_URL, data=payload).text 

@st.cache_data(ttl=60)
def load_data(mode):
    if mode == "maint": data = sheet_maint.get_all_records()
    elif mode == "users": data = sheet_users.get_all_records()
    else: data = sheet_demo.get_all_records()
    df = pd.DataFrame(data)
    if not df.empty and mode != "users":
        df = df.iloc[::-1].reset_index(drop=True)
    return df

st.markdown("""
<style>
.glide-card { background-color: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 12px; border-left: 6px solid #FFA726; }
.glide-title { font-size: 16px; font-weight: 700; color: #333333; margin-bottom: 4px; }
.glide-subtitle { font-size: 14px; color: #555555; margin-bottom: 10px; line-height: 1.4; }
.glide-tag { background-color: #FFF3E0; color: #E65100; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; margin-bottom: 6px; }
.glide-solution { font-size: 13px; color: #D84315; background-color: #FBE9E7; padding: 8px; border-radius: 6px; margin-top: 6px; }
.glide-img { width: 100%; border-radius: 8px; margin-top: 10px; border: 1px solid #eee; }
.calc-yellow { background-color: #FFF9C4; color: #333333; padding: 8px 12px; border-radius: 8px; border-left: 5px solid #FBC02D; font-weight: bold; margin-bottom: 10px; }
.calc-green { background-color: #E8F5E9; padding: 12px; border-radius: 8px; border-left: 6px solid #4CAF50; font-size: 18px; font-weight: bold; color: #2E7D32; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

if st.session_state.logged_in:
    now = datetime.now(tz_tw)
    if now - st.session_state.last_active > timedelta(minutes=15):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.warning("⏱️ 您已閒置超過 15 分鐘，為確保資安，系統已自動登出。請重新登入。")
        st.rerun()
    else:
        st.session_state.last_active = now

if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>🔐 設備維修與實驗知識庫</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #666;'>請輸入您的工號與密碼以登入系統</p>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        emp_id = st.text_input("工號 (EPM_ID)")
        password = st.text_input("密碼", type="password")
        submitted = st.form_submit_button("登入", use_container_width=True)
        
        if submitted:
            if not emp_id or not password:
                st.error("⚠️ 請輸入工號與密碼！")
            else:
                users_df = load_data("users")
                user_record = users_df[users_df['EPM_ID'].astype(str) == emp_id]
                if user_record.empty:
                    st.error("❌ 找不到此工號，請聯絡系統管理員。")
                else:
                    stored_hash = str(user_record.iloc[0]['Password_Hash'])
                    if hash_pw(password) == stored_hash:
                        st.session_state.logged_in = True
                        st.session_state.emp_id = emp_id
                        st.session_state.user_name = str(user_record.iloc[0]['Name'])
                        st.session_state.role = str(user_record.iloc[0]['Role'])
                        
                        if str(user_record.iloc[0]['Is_First_Login']).upper() == 'TRUE':
                            st.session_state.must_change_pw = True
                        else:
                            # 📌 強化版登入時間紀錄 (使用 update_acell 直接鎖定 F 欄)
                            cell = sheet_users.find(str(emp_id), in_column=1)
                            login_time = datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M:%S")
                            sheet_users.update_acell(f"F{cell.row}", login_time)
                            st.cache_data.clear()
                            
                        st.rerun()
                    else:
                        st.error("❌ 密碼錯誤，請重試！")

elif st.session_state.must_change_pw:
    st.warning("⚠️ 系統偵測到您為首次登入 (或密碼已被重置)，為確保資安，請強制修改密碼。")
    with st.form("change_pw_form"):
        new_pw = st.text_input("請輸入新密碼", type="password")
        confirm_pw = st.text_input("請再次輸入新密碼", type="password")
        pw_submitted = st.form_submit_button("確認修改並進入系統", use_container_width=True)
        
        if pw_submitted:
            if not new_pw or len(new_pw) < 4:
                st.error("⚠️ 密碼長度請至少輸入 4 個字元。")
            elif new_pw != confirm_pw:
                st.error("⚠️ 兩次密碼輸入不一致，請重新檢查！")
            else:
                with st.spinner("密碼加密更新中..."):
                    cell = sheet_users.find(str(st.session_state.emp_id), in_column=1)
                    sheet_users.update_acell(f"C{cell.row}", hash_pw(new_pw))
                    sheet_users.update_acell(f"E{cell.row}", "FALSE")
                    
                    # 📌 強化版登入時間紀錄 (強制改密碼後也寫入時間)
                    login_time = datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M:%S")
                    sheet_users.update_acell(f"F{cell.row}", login_time)
                    
                    st.cache_data.clear()
                    st.session_state.must_change_pw = False
                    st.success("✅ 密碼修改成功！正在為您導向主系統...")
                    st.rerun()
                    else:
    df_maint = load_data("maint")
    df_demo = load_data("demo")
    
    all_cust = []
    if not df_maint.empty: all_cust += df_maint['Customer'].astype(str).tolist()
    if not df_demo.empty: all_cust += df_demo['Customer'].astype(str).tolist()
    unique_cust = sorted(list(set([c.strip() for c in all_cust if c.strip() and c.strip() != 'nan'])))

    with st.sidebar:
        st.success(f"👤 歡迎登入，{st.session_state.user_name}！")
        st.markdown("### 🎛️ 系統模式切換")
        sys_options = ["🔧 現場維修系統", "🧪 DEMO 實驗紀錄", "🧮 產品厚度計算機"]
        if st.session_state.role == 'Admin': sys_options.append("👑 管理員後台")
        app_mode = st.radio("選擇要使用的系統：", sys_options, label_visibility="collapsed")
        
        if app_mode != "🧮 產品厚度計算機" and app_mode != "👑 管理員後台":
            st.write("---")
            st.markdown("### 📴 無塵室離線準備")
            if app_mode == "🔧 現場維修系統" and not df_maint.empty:
                st.download_button("📥 下載維修離線版", data=df_maint.to_csv(index=False).encode('utf-8-sig'), file_name=f"維修紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
            elif app_mode == "🧪 DEMO 實驗紀錄" and not df_demo.empty:
                st.download_button("📥 下載實驗離線版", data=df_demo.to_csv(index=False).encode('utf-8-sig'), file_name=f"實驗紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
        
        st.write("---")
        if st.button("🔄 重新整理最新資料", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        if st.button("🚪 登出系統", use_container_width=True):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

    if app_mode != "🧮 產品厚度計算機" and app_mode != "👑 管理員後台":
        col1, col2 = st.columns([1, 5])
        with col1:
            try: st.image("logo.png", width=80) 
            except: st.title("⚙️") 
        with col2:
            st.markdown(f"<h1 style='margin-top: -15px;'>{'設備維修知識庫' if app_mode == '🔧 現場維修系統' else 'DEMO 實驗資料庫'}</h1>", unsafe_allow_html=True)
            
        if st.session_state.success_msg:
            st.success(st.session_state.success_msg)
            st.session_state.success_msg = ""

    if app_mode == "🔧 現場維修系統":
        tab1, tab2, tab3, tab4 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄", "📊 數據分析", "✏️ 修改我的紀錄"])
        
        with tab1:
            search_keyword = st.text_input("🔍 全域搜尋 (例如: 日期, 廠區, 問題 ... 等 關鍵字)")
            filtered_df = df_maint.copy()
            if not filtered_df.empty:
                try: filtered_df['YearMonth'] = pd.to_datetime(filtered_df['Date']).dt.strftime('%y/%m')
                except: filtered_df['YearMonth'] = "未知時間"

                if search_keyword:
                    mask = pd.Series(False, index=filtered_df.index)
                    for col in filtered_df.columns: mask = mask | filtered_df[col].astype(str).str.contains(search_keyword, case=False, na=False)
                    filtered_df = filtered_df[mask]

                st.write("---")
                group_by_option = st.radio("🗂️ 選擇展開分類方式：", ["依建立年月", "依客戶與廠區", "依設備機型", "依設備部件"], horizontal=True)
                group_col_map = {"依建立年月": "YearMonth", "依客戶與廠區": "Customer", "依設備機型": "Machine_Model", "依設備部件": "Component"}
                group_col = group_col_map[group_by_option]
                
                st.caption(f"🔍 找到 {len(filtered_df)} 筆相關紀錄")
                custom_component_order = ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"]
                unique_groups = filtered_df[group_col].unique().tolist()
                if group_col == "Component": unique_groups.sort(key=lambda x: custom_component_order.index(x) if x in custom_component_order else 999)
                elif group_col == "YearMonth": unique_groups.sort(reverse=True) 
                else: unique_groups.sort(key=lambda x: str(x))

                for group_name in unique_groups:
                    display_name = group_name if str(group_name).strip() != "" else "未分類/未填寫"
                    group_data = filtered_df[filtered_df[group_col] == group_name]
                    with st.expander(f"📁 {display_name} (共 {len(group_data)} 筆)"):
                        for index, row in group_data.iterrows():
                            csv_data = pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig')
                            st.download_button(label=f"📥 下載此筆報告 ({row.get('Log_ID')})", data=csv_data, file_name=f"{row.get('Log_ID')}_Report.csv", mime="text/csv", key=f"dl_m_{row.get('Log_ID')}")
                            
                            photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                            val_issue = str(row.get('Issue_Desc', '')).replace('\r', '').replace('\n', '<br>')
                            val_solution = str(row.get('Solution', '')).replace('\r', '').replace('\n', '<br>')
                            st.markdown(f"""
    <div class="glide-card">
    <div class="glide-title">{row.get('Component', '')} <span style="font-size:12px; color:#999;">(單號: {row.get('Log_ID', '')})</span></div>
    <div class="glide-tag">📅 {row.get('Date', '')}</div>
    <div class="glide-tag">🏢 {row.get('Customer', '')}</div>
    <div class="glide-tag">⚙️ {row.get('Machine_Model', '')}</div>
    <div class="glide-tag" style="background-color:#E3F2FD; color:#1565C0;">👤 {row.get('Engineer', '')}</div>
    <div class="glide-subtitle"><b>狀況：</b><br>{val_issue}</div>
    <div class="glide-solution"><b>💡 解法：</b><br>{val_solution}</div>
    {photo_html}
    </div>
    """, unsafe_allow_html=True)
            else:
                st.warning("目前試算表中沒有資料喔！")

        with tab2:
            fk = st.session_state.form_key
            with st.form(f"maint_form_{fk}", clear_on_submit=True):
                st.subheader("📝 填寫維修紀錄")
                comp_options = ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"]
                
                input_date = st.date_input("日期", datetime.now(tz_tw).date(), key=f"m_date_{fk}")
                input_engineer = st.text_input("填單人員", value=st.session_state.user_name, disabled=True, key=f"m_eng_{fk}")
                
                c_cust1, c_cust2 = st.columns(2)
                with c_cust1: sel_cust = st.selectbox("選擇既有客戶廠區", [""] + unique_cust, key=f"m_cust_sel_{fk}")
                with c_cust2: new_cust = st.text_input("或自填新廠區", placeholder="若左側無相符請填此", key=f"m_cust_new_{fk}")
                
                input_machine = st.selectbox("設備機型", ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"], index=None, key=f"m_mach_{fk}")
                input_component = st.selectbox("異常部件", comp_options, index=None, key=f"m_comp_{fk}")
                input_issue = st.text_area("問題描述", key=f"m_iss_{fk}")
                input_solution = st.text_area("解決方案", key=f"m_sol_{fk}")
                upload_file = st.file_uploader("🖼️ 附加現場照片", type=['jpg', 'png', 'jpeg'], key=f"m_photo_{fk}")
                
                st.write("---")
                maint_msg = st.empty()
                if st.form_submit_button("送出維修紀錄", key=f"btn_m_{fk}"):
                    final_customer = new_cust.strip() if new_cust.strip() else sel_cust
                    if not all([final_customer, input_machine, input_component, input_issue, input_solution]):
                        maint_msg.error("⚠️ 請確認所有必填欄位都已填寫！")
                    else:
                        with st.spinner("寫入中..."):
                            log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M")
                            photo_url = upload_image(upload_file, f"{log_id}.jpg") if upload_file else ""
                            sheet_maint.append_row([log_id, input_date.strftime("%Y-%m-%d"), input_engineer, final_customer, input_machine, input_component, input_issue, input_solution, photo_url])
                            st.cache_data.clear()
                            maint_msg.success(f"✅ 成功寫入資料庫！單號：{log_id}")

        with tab3:
            st.subheader("📈 維修數據統計看板")
            if not df_maint.empty:
                col1, col2, col3 = st.columns(3)
                with col1: st.metric("累積維修總件數", f"{len(df_maint)} 件")
                with col2: st.metric("本月新增件數", f"{df_maint[df_maint['Date'].str.startswith(datetime.now(tz_tw).strftime('%Y-%m'), na=False)].shape[0]} 件")
                with col3: st.metric("涵蓋機型數量", f"{df_maint['Machine_Model'].nunique()} 種")
            else:
                st.info("目前系統中還沒有資料！")

        with tab4:
            st.subheader("✏️ 修改我的維修紀錄")
            if st.session_state.role == 'Admin':
                my_df = df_maint.copy()
            else:
                my_df = df_maint[df_maint['Engineer'].astype(str).str.strip() == st.session_state.user_name]
            
            if my_df.empty:
                st.warning("您目前還沒有建立過任何維修紀錄喔！")
            else:
                options = [""] + [f"{r['Log_ID']} (日期: {r['Date']} | 客戶: {r['Customer']} | 機台: {r['Machine_Model']})" for idx, r in my_df.iterrows()]
                selected_m = st.selectbox("🔍 請選擇要修改的維修單 (支援關鍵字搜尋)", options, key="select_edit_m")
                
                if selected_m:
                    edit_m_id = selected_m.split(" ")[0]
                    row_dict = df_maint[df_maint['Log_ID'] == edit_m_id].iloc[0].to_dict()
                    
                    st.success(f"✅ 成功載入單號：{edit_m_id} (若不需修改照片請留空)")
                    with st.form(f"edit_m_form_{edit_m_id}", clear_on_submit=True):
                        try: old_date = datetime.strptime(str(row_dict.get('Date', '')), "%Y-%m-%d").date()
                        except: old_date = datetime.now(tz_tw).date()
                        
                        e_date = st.date_input("日期", value=old_date, key=f"em_date_{edit_m_id}")
                        e_engineer = st.text_input("填單人員", value=str(row_dict.get('Engineer', '')), disabled=True, key=f"em_eng_{edit_m_id}")
                        e_customer = st.text_input("客戶與廠區", value=str(row_dict.get('Customer', '')), key=f"em_cust_{edit_m_id}")
                        
                        mach_opts = ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"]
                        old_mach = str(row_dict.get('Machine_Model', ''))
                        m_idx = mach_opts.index(old_mach) if old_mach in mach_opts else None
                        e_machine = st.selectbox("設備機型", mach_opts, index=m_idx, key=f"em_mach_{edit_m_id}")
                        
                        comp_options = ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"]
                        old_comp = str(row_dict.get('Component', ''))
                        c_idx = comp_options.index(old_comp) if old_comp in comp_options else None
                        e_component = st.selectbox("異常部件", comp_options, index=c_idx, key=f"em_comp_{edit_m_id}")
                        
                        e_issue = st.text_area("問題描述", value=str(row_dict.get('Issue_Desc', '')), key=f"em_iss_{edit_m_id}")
                        e_solution = st.text_area("解決方案", value=str(row_dict.get('Solution', '')), key=f"em_sol_{edit_m_id}")
                        e_upload = st.file_uploader("🖼️ 更新現場照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"em_photo_{edit_m_id}")
                        
                        st.write("---")
                        edit_m_msg = st.empty()
                        if st.form_submit_button("💾 覆蓋更新紀錄", key=f"btn_em_{edit_m_id}"):
                            with st.spinner("更新雲端資料庫中..."):
                                new_photo_url = upload_image(e_upload, f"{edit_m_id}_edit.jpg") if e_upload else str(row_dict.get('Photo_URL', ''))
                                new_m_row = [edit_m_id, e_date.strftime("%Y-%m-%d"), e_engineer, e_customer, e_machine, e_component, e_issue, e_solution, new_photo_url]
                                cell = sheet_maint.find(edit_m_id, in_column=1)
                                sheet_maint.update(values=[new_m_row], range_name=f"A{cell.row}:I{cell.row}")
                                st.cache_data.clear()
                                edit_m_msg.success(f"✅ 單號 {edit_m_id} 更新成功！")
                                elif app_mode == "🧪 DEMO 實驗紀錄":
        tab_d1, tab_d2, tab_d3, tab_d4 = st.tabs(["🔍 參數查詢", "➕ 新增紀錄 (NT+CVP)", "➕ 新增紀錄 (V-160)", "✏️ 修改我的紀錄"])

        with tab_d1:
            c_search, c_filter1, c_filter2 = st.columns([2, 1, 1])
            with c_search: search_kw_demo = st.text_input("🔍 全域搜尋 (關鍵字)")
            sub_types = ["全部"]
            film_mats = ["全部"]
            if not df_demo.empty:
                if 'Substrate_Type' in df_demo.columns: sub_types += sorted([str(x) for x in df_demo['Substrate_Type'].unique() if str(x).strip() and str(x) != 'nan'])
                if 'Film_Material' in df_demo.columns: film_mats += sorted([str(x) for x in df_demo['Film_Material'].unique() if str(x).strip() and str(x) != 'nan'])
            with c_filter1: filter_sub = st.selectbox("🗂️ 依板材篩選", sub_types)
            with c_filter2: filter_film = st.selectbox("🗂️ 依膜材篩選", film_mats)

            filtered_demo = df_demo.copy()
            if not filtered_demo.empty:
                try: filtered_demo['YearMonth'] = pd.to_datetime(filtered_demo['Date']).dt.strftime('%y/%m')
                except: filtered_demo['YearMonth'] = "未知時間"

                if search_kw_demo:
                    mask = pd.Series(False, index=filtered_demo.index)
                    for col in filtered_demo.columns: mask = mask | filtered_demo[col].astype(str).str.contains(search_kw_demo, case=False, na=False)
                    filtered_demo = filtered_demo[mask]
                if filter_sub != "全部": filtered_demo = filtered_demo[filtered_demo['Substrate_Type'].astype(str) == filter_sub]
                if filter_film != "全部": filtered_demo = filtered_demo[filtered_demo['Film_Material'].astype(str) == filter_film]
                
                st.write("---")
                group_by_demo = st.radio("🗂️ 選擇展開分類方式：", ["依建立年月", "依客戶名稱", "依測試機台"], horizontal=True)
                group_col_demo_map = {"依建立年月": "YearMonth", "依客戶名稱": "Customer", "依測試機台": "Equipment"}
                group_col_d = group_col_demo_map[group_by_demo]

                st.caption(f"🔍 找到 {len(filtered_demo)} 筆實驗紀錄")
                unique_groups_d = filtered_demo[group_col_d].unique().tolist()
                if group_col_d == "YearMonth": unique_groups_d.sort(reverse=True) 
                else: unique_groups_d.sort(key=lambda x: str(x))

                for group_name in unique_groups_d:
                    display_name = group_name if str(group_name).strip() != "" else "未分類/未填寫"
                    group_data = filtered_demo[filtered_demo[group_col_d] == group_name]
                    with st.expander(f"📁 {display_name} (共 {len(group_data)} 筆)"):
                        for index, row in group_data.iterrows():
                            csv_data = pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig')
                            st.download_button(label=f"📥 下載此筆報告 ({row.get('Log_ID')})", data=csv_data, file_name=f"{row.get('Log_ID')}_Report.csv", mime="text/csv", key=f"dl_d_{row.get('Log_ID')}")
                            
                            photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                            equip_name = str(row.get('Equipment', ''))
                            
                            sub_dict = {"板材類型": str(row.get('Substrate_Type', '')).strip(), "基板尺寸": str(row.get('Substrate_Size', '')).strip(), "膜材種類": str(row.get('Film_Material', '')).strip(), "膜材型號 / 厚度": str(row.get('Film_Model', '')).strip()}
                            html_sub = format_params_html(pack_params(sub_dict))
                            
                            html_pre = format_params_html(row.get('Pre_Lam', ''))
                            html_l1 = format_params_html(row.get('Lam_1st', ''))
                            html_l2 = format_params_html(row.get('Lam_2nd', ''))
                            html_l3 = format_params_html(row.get('Lam_3rd', ''))
                            
                            blocks = []
                            if "V-160" in equip_name.upper() or "V160" in equip_name.upper():
                                if html_l1: blocks.append(f"<b>🔹 V-160 參數</b><br>{html_l1}")
                            else:
                                if html_pre: blocks.append(f"<b>🔹 預貼機參數</b><br>{html_pre}")
                                if html_l1: blocks.append(f"<b>🔹 1st 壓模</b><br>{html_l1}")
                                if html_l2: blocks.append(f"<b>🔹 2nd 壓模</b><br>{html_l2}")
                                if html_l3: blocks.append(f"<b>🔹 3rd 壓模</b><br>{html_l3}")
                            
                            params_block = f"<div style='background-color:#F9F9F9; padding:10px; border-radius:8px; margin-bottom:10px; font-size:13px; color:#555;'>{'<br><br>'.join(blocks)}</div>" if blocks else ""
                            sub_block = f"<div class='glide-subtitle'><b>基材/膜材</b><br>{html_sub}</div>" if html_sub else ""
                            eval_result = str(row.get('Self_Eval', '未評估'))
                            if not eval_result or eval_result == "nan": eval_result = "未評估"

                            st.markdown(f"""
            <div class="glide-card">
            <div class="glide-title">🧪 測試機台: {equip_name if equip_name else '未填寫'} <span style="font-size:12px; color:#999;">(單號: {row.get('Log_ID', '')})</span></div>
            <div class="glide-tag">📅 {row.get('Date', '')}</div>
            <div class="glide-tag">🏢 {row.get('Customer', '')}</div>
            <div class="glide-tag" style="background-color:#E3F2FD; color:#1565C0;">👤 操作: {row.get('Operator', '')}</div>
            <div class="glide-tag">📦 數量: {row.get('Qty', '')}</div>
            <div class="glide-tag" style="background-color: #E8F5E9; color: #2E7D32;">📊 自評: {eval_result}</div>
            {sub_block}
            {params_block}
            <div class="glide-subtitle"><b>📝 備註與異常</b><br>{str(row.get('Remarks', '無')).replace('\n', '<br>')}</div>
            <div class="glide-solution"><b>🗣️ 客戶反饋</b><br>{str(row.get('Feedback', '無')).replace('\n', '<br>')}</div>
            {photo_html}
            </div>
            """, unsafe_allow_html=True)
            else:
                st.info("目前還沒有 DEMO 實驗紀錄，趕快去新增一筆吧！")

        with tab_d2:
            fk = st.session_state.form_key
            st.subheader("🧪 填寫實驗紀錄 (NT+CVP)")
            
            clone_nt_dict = {}
            if "load_nt_key" not in st.session_state: st.session_state.load_nt_key = 0

            if not df_demo.empty:
                history_opts_nt = [""] + [f"{r['Log_ID']} (客戶:{r['Customer']} | 機台:{r['Equipment']})" for idx, r in df_demo.iterrows() if "V-160" not in str(r['Equipment']).upper() and "V160" not in str(r['Equipment']).upper()]
                clone_sel_nt = st.selectbox("⚡ 一鍵帶入歷史參數 (選擇後自動填滿下方欄位)", history_opts_nt, key=f"clone_nt_sel_{fk}")
                
                if clone_sel_nt:
                    if st.session_state.get(f"last_clone_nt_{fk}") != clone_sel_nt:
                        st.session_state.load_nt_key += 1
                        st.session_state[f"last_clone_nt_{fk}"] = clone_sel_nt
                    clone_id = clone_sel_nt.split(" ")[0]
                    match_row = df_demo[df_demo['Log_ID'].astype(str).str.strip() == clone_id.strip()]
                    if not match_row.empty: clone_nt_dict = match_row.iloc[0].to_dict()
                else:
                    st.session_state[f"last_clone_nt_{fk}"] = ""

            lk_nt = st.session_state.load_nt_key
            
            with st.form(f"demo_form_{fk}", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1: input_d_date = st.date_input("測試日期", datetime.now(tz_tw).date(), key=f"d_date_{fk}")
                
                c_cust1, c_cust2 = st.columns(2)
                old_cust = str(clone_nt_dict.get('Customer', ''))
                c_idx = unique_cust.index(old_cust) + 1 if old_cust in unique_cust else 0
                with c_cust1: sel_cust = st.selectbox("選擇既有客戶", [""] + unique_cust, index=c_idx, key=f"d_cust_sel_{fk}_{lk_nt}")
                with c_cust2: new_cust = st.text_input("或自填新客戶", value=old_cust if old_cust and c_idx == 0 else "", placeholder="若無相符請填此", key=f"d_cust_new_{fk}_{lk_nt}")
                
                c4, c5 = st.columns(2)
                with c4: input_d_operator = st.text_input("操作人", value=st.session_state.user_name, disabled=True, key=f"d_oper_{fk}_{lk_nt}")
                with c5: input_d_equip = st.text_input("設備類型 (如: CVP-1600SP)", value=str(clone_nt_dict.get('Equipment', '')), key=f"d_equip_{fk}_{lk_nt}")
                
                with st.expander("📍 基材與膜材資訊", expanded=True):
                    c_s1, c_s2 = st.columns(2)
                    st_opts = ["", "PCB", "Wafer", "Glass", "其他"]
                    c_st = str(clone_nt_dict.get('Substrate_Type', ''))
                    st_idx = st_opts.index(c_st) if c_st in st_opts else (4 if c_st else 0)
                    with c_s1: 
                        input_d_sub_t = st.selectbox("板材類型", st_opts, index=st_idx, key=f"s_t_{fk}_{lk_nt}")
                        input_d_sub_t_other = st.text_input("自填板材", value=c_st if st_idx == 4 else "", label_visibility="collapsed", placeholder="若選其他請在此填寫", key=f"s_to_{fk}_{lk_nt}")
                        input_d_sub_size = st.text_input("基板尺寸與厚度", value=str(clone_nt_dict.get('Substrate_Size', '')), key=f"s_d_{fk}_{lk_nt}")
                    
                    fm_opts = ["", "ABF", "DAF", "NCF", "PI", "其他"]
                    c_fm = str(clone_nt_dict.get('Film_Material', ''))
                    fm_idx = fm_opts.index(c_fm) if c_fm in fm_opts else (5 if c_fm else 0)
                    with c_s2: 
                        input_d_film_m = st.selectbox("膜材種類", fm_opts, index=fm_idx, key=f"f_m_{fk}_{lk_nt}")
                        input_d_film_m_other = st.text_input("自填膜材", value=c_fm if fm_idx == 5 else "", label_visibility="collapsed", placeholder="若選其他請在此填寫", key=f"f_mo_{fk}_{lk_nt}")
                        input_d_film_model = st.text_input("膜材型號 / 厚度", value=str(clone_nt_dict.get('Film_Model', '')), key=f"f_mod_{fk}_{lk_nt}")
                
                st.write("---")
                pre_defs = unpack_params(clone_nt_dict.get('Pre_Lam', ''))
                with st.expander("📍 預貼機參數"):
                    c_p1, c_p2 = st.columns(2)
                    with c_p1: pre_t = st.text_input("預貼溫度 (℃)", value=pre_defs.get("預貼溫度 (℃)", ""), key=f"p_t_{fk}_{lk_nt}")
                    with c_p2: pre_s = st.text_input("預貼速度 (m/min)", value=pre_defs.get("預貼速度 (m/min)", ""), key=f"p_s_{fk}_{lk_nt}")
                    c_p3, c_p4 = st.columns(2)
                    with c_p3: pre_p = st.text_input("預貼壓力 (MPa)", value=pre_defs.get("預貼壓力 (MPa)", ""), key=f"p_p_{fk}_{lk_nt}")
                    with c_p4: pre_m = st.text_input("前後留邊量 (前mm / 後mm)", value=pre_defs.get("前後留邊量", ""), key=f"p_m_{fk}_{lk_nt}")
                        
                lam1_dict = render_lam_inputs("1st 壓模機", "l1", f"{fk}_{lk_nt}", unpack_params(clone_nt_dict.get('Lam_1st', '')))
                lam2_dict = render_lam_inputs("2nd 壓模機", "l2", f"{fk}_{lk_nt}", unpack_params(clone_nt_dict.get('Lam_2nd', '')))
                lam3_dict = render_lam3_inputs("3rd 壓模機", "l3", f"{fk}_{lk_nt}", unpack_params(clone_nt_dict.get('Lam_3rd', '')))
                    
                st.write("---")
                c_q1, c_q2 = st.columns(2)
                with c_q1: input_d_qty = st.text_input("壓合數量 (片/次)", value=str(clone_nt_dict.get('Qty', '')), key=f"d_qty_{fk}_{lk_nt}")
                
                eval_opts = ["⚪ 尚未評估", "🟢 佳 (參數可參考)", "🟡 普通 (需微調)", "🔴 差 (不建議使用)"]
                old_eval = str(clone_nt_dict.get('Self_Eval', ''))
                e_idx = eval_opts.index(old_eval) if old_eval in eval_opts else 0
                with c_q2: input_d_eval = st.selectbox("內部自評結果", eval_opts, index=e_idx, key=f"d_eval_{fk}_{lk_nt}")
                
                input_d_remark = st.text_area("備註 (測試變動說明、具體異常)", value=str(clone_nt_dict.get('Remarks', '')), key=f"d_rmk_{fk}_{lk_nt}")
                input_d_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)", value=str(clone_nt_dict.get('Feedback', '')), key=f"d_fb_{fk}_{lk_nt}")
                upload_d_file = st.file_uploader("🖼️ 附加測試結果照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"d_photo_{fk}_{lk_nt}")
                
                st.write("---")
                demo_msg = st.empty()
                if st.form_submit_button("送出實驗紀錄", key=f"btn_d_{fk}"):
                    final_customer = new_cust.strip() if new_cust.strip() else sel_cust
                    if not all([final_customer, input_d_equip]):
                        demo_msg.error("⚠️ 請至少填寫客戶名稱與設備類型！")
                    else:
                        with st.spinner("打包參數並寫入雲端中..."):
                            final_d_sub_t = input_d_sub_t_other if input_d_sub_t == "其他" else input_d_sub_t
                            final_d_film_m = input_d_film_m_other if input_d_film_m == "其他" else input_d_film_m
                            input_d_pre = pack_params({"預貼溫度 (℃)": pre_t, "預貼壓力 (MPa)": pre_p, "預貼速度 (m/min)": pre_s, "前後留邊量": pre_m})
                            input_d_1st = pack_params(lam1_dict)
                            input_d_2nd = pack_params(lam2_dict)
                            input_d_3rd = pack_params(lam3_dict)
                            log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                            photo_url = upload_image(upload_d_file, f"{log_id}.jpg") if upload_d_file else ""
                            new_demo_row = [log_id, input_d_date.strftime("%Y-%m-%d"), input_d_operator, final_customer, input_d_equip, final_d_sub_t, input_d_sub_size, final_d_film_m, input_d_film_model, input_d_pre, input_d_1st, input_d_2nd, input_d_3rd, input_d_qty, input_d_eval, input_d_remark, input_d_feedback, photo_url]
                            sheet_demo.append_row(new_demo_row)
                            st.cache_data.clear()
                            
                            st.session_state[f"clone_nt_sel_{fk}"] = ""
                            st.session_state[f"last_clone_nt_{fk}"] = ""
                            st.session_state.load_nt_key += 1
                            demo_msg.success(f"✅ 成功寫入 DEMO 紀錄！單號：{log_id}")

        with tab_d3:
            fk = st.session_state.form_key
            st.subheader("🧪 填寫實驗紀錄 (V-160)")
            
            clone_v_dict = {}
            if "load_v_key" not in st.session_state: st.session_state.load_v_key = 0

            if not df_demo.empty:
                history_opts_v = [""] + [f"{r['Log_ID']} (客戶:{r['Customer']} | 機台:{r['Equipment']})" for idx, r in df_demo.iterrows() if "V-160" in str(r['Equipment']).upper() or "V160" in str(r['Equipment']).upper()]
                clone_sel_v = st.selectbox("⚡ 一鍵帶入歷史參數 (選擇後自動填滿下方欄位)", history_opts_v, key=f"clone_v_sel_{fk}")
                
                if clone_sel_v:
                    if st.session_state.get(f"last_clone_v_{fk}") != clone_sel_v:
                        st.session_state.load_v_key += 1
                        st.session_state[f"last_clone_v_{fk}"] = clone_sel_v
                    clone_id = clone_sel_v.split(" ")[0]
                    match_row = df_demo[df_demo['Log_ID'].astype(str).str.strip() == clone_id.strip()]
                    if not match_row.empty: clone_v_dict = match_row.iloc[0].to_dict()
                else:
                    st.session_state[f"last_clone_v_{fk}"] = ""

            lk_v = st.session_state.load_v_key

            with st.form(f"v160_form_{fk}", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1: input_v_date = st.date_input("測試日期", datetime.now(tz_tw).date(), key=f"v_date_{fk}")
                
                c_cust1, c_cust2 = st.columns(2)
                old_cust_v = str(clone_v_dict.get('Customer', ''))
                c_idx_v = unique_cust.index(old_cust_v) + 1 if old_cust_v in unique_cust else 0
                with c_cust1: sel_cust_v = st.selectbox("選擇既有客戶", [""] + unique_cust, index=c_idx_v, key=f"v_cust_sel_{fk}_{lk_v}")
                with c_cust2: new_cust_v = st.text_input("或自填新客戶", value=old_cust_v if old_cust_v and c_idx_v == 0 else "", placeholder="若無相符請填此", key=f"v_cust_new_{fk}_{lk_v}")
                
                c4, c5 = st.columns(2)
                with c4: input_v_operator = st.text_input("操作人", value=st.session_state.user_name, disabled=True, key=f"v_oper_{fk}_{lk_v}")
                with c5: input_v_equip = st.text_input("設備類型", value="V-160", disabled=True, key=f"v_equip_{fk}_{lk_v}") 
                
                with st.expander("📍 基材與膜材資訊", expanded=True):
                    c_vs1, c_vs2 = st.columns(2)
                    st_opts = ["", "PCB", "Wafer", "Glass", "其他"]
                    c_st_v = str(clone_v_dict.get('Substrate_Type', ''))
                    st_idx_v = st_opts.index(c_st_v) if c_st_v in st_opts else (4 if c_st_v else 0)
                    with c_vs1: 
                        input_v_sub_t = st.selectbox("板材類型", st_opts, index=st_idx_v, key=f"v_s_t_{fk}_{lk_v}")
                        input_v_sub_t_other = st.text_input("自填板材", value=c_st_v if st_idx_v == 4 else "", label_visibility="collapsed", placeholder="若選其他請在此填寫", key=f"v_s_to_{fk}_{lk_v}")
                        input_v_sub_size = st.text_input("基板尺寸與厚度", value=str(clone_v_dict.get('Substrate_Size', '')), key=f"v_s_d_{fk}_{lk_v}")
                    
                    fm_opts = ["", "ABF", "DAF", "NCF", "PI", "其他"]
                    c_fm_v = str(clone_v_dict.get('Film_Material', ''))
                    fm_idx_v = fm_opts.index(c_fm_v) if c_fm_v in fm_opts else (5 if c_fm_v else 0)
                    with c_vs2: 
                        input_v_film_m = st.selectbox("膜材種類", fm_opts, index=fm_idx_v, key=f"v_f_m_{fk}_{lk_v}")
                        input_v_film_m_other = st.text_input("自填膜材", value=c_fm_v if fm_idx_v == 5 else "", label_visibility="collapsed", placeholder="若選其他請在此填寫", key=f"v_f_mo_{fk}_{lk_v}")
                        input_v_film_model = st.text_input("膜材型號 / 厚度", value=str(clone_v_dict.get('Film_Model', '')), key=f"v_f_mod_{fk}_{lk_v}")
                
                st.write("---")
                v_defs = unpack_params(clone_v_dict.get('Lam_1st', ''))
                with st.expander("📍 V-160 參數"):
                    c_v1, c_v2 = st.columns(2)
                    v_modes = ["", "上", "下", "上下"]
                    old_vmode = v_defs.get("加壓模式", "")
                    vm_idx = v_modes.index(old_vmode) if old_vmode in v_modes else 0
                    with c_v1: v_mode = st.selectbox("加壓模式", v_modes, index=vm_idx, key=f"v_mode_{fk}_{lk_v}")
                    with c_v2: v_tv = st.text_input("下真空時間 (sec)", value=v_defs.get("下真空時間 (sec)", ""), key=f"v_tv_{fk}_{lk_v}")
                    c_v3, c_v4 = st.columns(2)
                    with c_v3: v_tt = st.text_input("上溫度 (℃)", value=v_defs.get("上溫度 (℃)", ""), key=f"v_tt_{fk}_{lk_v}")
                    with c_v4: v_tb = st.text_input("下溫度 (℃)", value=v_defs.get("下溫度 (℃)", ""), key=f"v_tb_{fk}_{lk_v}")
                    st.write("---")
                    v_tdrop_t = st.text_input("上硅膠墊垂落時間 (sec)", value=v_defs.get("上硅膠墊垂落時間 (sec)", ""), key=f"v_tdrop_t_{fk}_{lk_v}")
                    c_v5, c_v6 = st.columns(2)
                    with c_v5: v_pt = st.text_input("上氣囊加壓壓力 (kgf/cm²)", value=v_defs.get("上氣囊加壓壓力 (kgf/cm²)", ""), key=f"v_pt_{fk}_{lk_v}")
                    with c_v6: v_tpt = st.text_input("上氣囊加壓時間 (sec)", value=v_defs.get("上氣囊加壓時間 (sec)", ""), key=f"v_tpt_{fk}_{lk_v}")
                    st.write("---")
                    v_dly_b = st.text_input("下加壓延遲時間 (sec)", value=v_defs.get("下加壓延遲時間 (sec)", ""), key=f"v_dly_b_{fk}_{lk_v}")
                    v_tdrop_b = st.text_input("下硅膠墊垂落時間 (sec)", value=v_defs.get("下硅膠墊垂落時間 (sec)", ""), key=f"v_tdrop_b_{fk}_{lk_v}")
                    c_v7, c_v8 = st.columns(2)
                    with c_v7: v_pb = st.text_input("下加壓壓力 (kgf/cm²)", value=v_defs.get("下加壓壓力 (kgf/cm²)", ""), key=f"v_pb_{fk}_{lk_v}")
                    with c_v8: v_tpb = st.text_input("下加壓時間 (sec)", value=v_defs.get("下加壓時間 (sec)", ""), key=f"v_tpb_{fk}_{lk_v}")
                    
                st.write("---")
                c_vq1, c_vq2 = st.columns(2)
                with c_vq1: input_v_qty = st.text_input("壓合數量 (片/次)", value=str(clone_v_dict.get('Qty', '')), key=f"v_qty_{fk}_{lk_v}")
                
                eval_opts = ["⚪ 尚未評估", "🟢 佳 (參數可參考)", "🟡 普通 (需微調)", "🔴 差 (不建議使用)"]
                old_eval = str(clone_v_dict.get('Self_Eval', ''))
                e_idx = eval_opts.index(old_eval) if old_eval in eval_opts else 0
                with c_vq2: input_v_eval = st.selectbox("內部自評結果", eval_opts, index=e_idx, key=f"v_eval_{fk}_{lk_v}")
                
                input_v_remark = st.text_area("備註 (測試變動說明、具體異常)", value=str(clone_v_dict.get('Remarks', '')), key=f"v_rmk_{fk}_{lk_v}")
                input_v_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)", value=str(clone_v_dict.get('Feedback', '')), key=f"v_fb_{fk}_{lk_v}")
                upload_v_file = st.file_uploader("🖼️ 附加測試結果照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"v_photo_{fk}_{lk_v}")
                
                st.write("---")
                v160_msg = st.empty()
                if st.form_submit_button("送出實驗紀錄", key=f"btn_v_{fk}"):
                    final_customer = new_cust_v.strip() if new_cust_v.strip() else sel_cust_v
                    if not all([final_customer]):
                        v160_msg.error("⚠️ 請至少填寫客戶名稱！")
                    else:
                        with st.spinner("打包 V-160 參數並寫入雲端中..."):
                            final_v_sub_t = input_v_sub_t_other if input_v_sub_t == "其他" else input_v_sub_t
                            final_v_film_m = input_v_film_m_other if input_v_film_m == "其他" else input_v_film_m
                            v160_dict = {
                                "加壓模式": v_mode, "下真空時間 (sec)": v_tv, "上溫度 (℃)": v_tt, "下溫度 (℃)": v_tb,
                                "上硅膠墊垂落時間 (sec)": v_tdrop_t, "上氣囊加壓壓力 (kgf/cm²)": v_pt, "上氣囊加壓時間 (sec)": v_tpt,
                                "下加壓延遲時間 (sec)": v_dly_b, "下硅膠墊垂落時間 (sec)": v_tdrop_b,
                                "下加壓壓力 (kgf/cm²)": v_pb, "下加壓時間 (sec)": v_tpb
                            }
                            input_v_params = pack_params(v160_dict)
                            log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                            photo_url = upload_image(upload_v_file, f"{log_id}.jpg") if upload_v_file else ""
                            new_v160_row = [log_id, input_v_date.strftime("%Y-%m-%d"), input_v_operator, final_customer, input_v_equip, final_v_sub_t, input_v_sub_size, final_v_film_m, input_v_film_model, "無", input_v_params, "無", "無", input_v_qty, input_v_eval, input_v_remark, input_v_feedback, photo_url]
                            sheet_demo.append_row(new_v160_row)
                            st.cache_data.clear()
                            
                            st.session_state[f"clone_v_sel_{fk}"] = ""
                            st.session_state[f"last_clone_v_{fk}"] = ""
                            st.session_state.load_v_key += 1
                            v160_msg.success(f"✅ 成功寫入 V-160 紀錄！單號：{log_id}")

        with tab_d4:
            st.subheader("✏️ 修改我的實驗紀錄")
            if st.session_state.role == 'Admin':
                my_df_d = df_demo.copy()
            else:
                my_df_d = df_demo[df_demo['Operator'].astype(str).str.strip() == st.session_state.user_name]
            
            if my_df_d.empty:
                st.warning("您目前還沒有建立過任何實驗紀錄喔！")
            else:
                options_d = [""] + [f"{r['Log_ID']} (日期: {r['Date']} | 客戶: {r['Customer']} | 機台: {r['Equipment']})" for idx, r in my_df_d.iterrows()]
                selected_d = st.selectbox("🔍 請選擇要修改的實驗單 (支援關鍵字搜尋)", options_d, key="select_edit_d")
                
                if selected_d:
                    edit_d_id = selected_d.split(" ")[0]
                    row_dict = df_demo[df_demo['Log_ID'] == edit_d_id].iloc[0].to_dict()
                    
                    st.success(f"✅ 成功載入實驗單號：{edit_d_id} (若不需修改照片請留空)")
                    with st.form(f"edit_d_form_{edit_d_id}", clear_on_submit=True):
                        try: old_date = datetime.strptime(str(row_dict.get('Date', '')), "%Y-%m-%d").date()
                        except: old_date = datetime.now(tz_tw).date()
                        
                        c1, c2 = st.columns(2)
                        ed_date = st.date_input("測試日期", value=old_date, key=f"ed_date_{edit_d_id}")
                        ed_customer = st.text_input("客戶名稱", value=str(row_dict.get('Customer', '')), key=f"ed_cust_{edit_d_id}")
                        
                        c3, c4 = st.columns(2)
                        with c3: ed_operator = st.text_input("操作人", value=str(row_dict.get('Operator', '')), disabled=True, key=f"ed_oper_{edit_d_id}")
                        with c4: ed_equip = st.text_input("設備類型", value=str(row_dict.get('Equipment', '')), key=f"ed_equip_{edit_d_id}")
                        
                        with st.expander("📍 修改：基材與膜材資訊", expanded=True):
                            c_s1, c_s2 = st.columns(2)
                            old_st = str(row_dict.get('Substrate_Type', ''))
                            st_opts = ["", "PCB", "Wafer", "Glass", "其他"]
                            st_idx = st_opts.index(old_st) if old_st in st_opts else 4
                            with c_s1: 
                                ed_sub_t = st.selectbox("板材類型", st_opts, index=st_idx, key=f"ed_st_{edit_d_id}")
                                ed_sub_t_other = st.text_input("自填板材", value=old_st if st_idx == 4 else "", label_visibility="collapsed", key=f"ed_sto_{edit_d_id}")
                                ed_sub_size = st.text_input("基板尺寸與厚度", value=str(row_dict.get('Substrate_Size', '')), key=f"ed_sd_{edit_d_id}")
                            
                            old_fm = str(row_dict.get('Film_Material', ''))
                            fm_opts = ["", "ABF", "DAF", "NCF", "PI", "其他"]
                            fm_idx = fm_opts.index(old_fm) if old_fm in fm_opts else 5
                            with c_s2: 
                                ed_film_m = st.selectbox("膜材種類", fm_opts, index=fm_idx, key=f"ed_fm_{edit_d_id}")
                                ed_film_m_other = st.text_input("自填膜材", value=old_fm if fm_idx == 5 else "", label_visibility="collapsed", key=f"ed_fmo_{edit_d_id}")
                                ed_film_model = st.text_input("膜材型號 / 厚度", value=str(row_dict.get('Film_Model', '')), key=f"ed_fmod_{edit_d_id}")
                        
                        st.write("---")
                        is_v160 = "V-160" in str(row_dict.get('Equipment', '')).upper() or "V160" in str(row_dict.get('Equipment', '')).upper()
                        
                        if is_v160:
                            v_defs = unpack_params(row_dict.get('Lam_1st', ''))
                            with st.expander("📍 V-160 參數", expanded=True):
                                c_v1, c_v2 = st.columns(2)
                                v_modes = ["", "上", "下", "上下"]
                                old_vmode = v_defs.get("加壓模式", "")
                                vm_idx = v_modes.index(old_vmode) if old_vmode in v_modes else 0
                                with c_v1: ed_v_mode = st.selectbox("加壓模式", v_modes, index=vm_idx, key=f"ed_vm_{edit_d_id}")
                                with c_v2: ed_v_tv = st.text_input("下真空時間 (sec)", value=v_defs.get("下真空時間 (sec)", ""), key=f"ed_v_tv_{edit_d_id}")
                                c_v3, c_v4 = st.columns(2)
                                with c_v3: ed_v_tt = st.text_input("上溫度 (℃)", value=v_defs.get("上溫度 (℃)", ""), key=f"ed_v_tt_{edit_d_id}")
                                with c_v4: ed_v_tb = st.text_input("下溫度 (℃)", value=v_defs.get("下溫度 (℃)", ""), key=f"ed_v_tb_{edit_d_id}")
                                st.write("---")
                                ed_v_tdrop_t = st.text_input("上硅膠墊垂落時間 (sec)", value=v_defs.get("上硅膠墊垂落時間 (sec)", ""), key=f"ed_v_tdt_{edit_d_id}")
                                c_v5, c_v6 = st.columns(2)
                                with c_v5: ed_v_pt = st.text_input("上氣囊加壓壓力 (kgf/cm²)", value=v_defs.get("上氣囊加壓壓力 (kgf/cm²)", ""), key=f"ed_v_pt_{edit_d_id}")
                                with c_v6: ed_v_tpt = st.text_input("上氣囊加壓時間 (sec)", value=v_defs.get("上氣囊加壓時間 (sec)", ""), key=f"ed_v_tpt_{edit_d_id}")
                                st.write("---")
                                ed_v_dly_b = st.text_input("下加壓延遲時間 (sec)", value=v_defs.get("下加壓延遲時間 (sec)", ""), key=f"ed_v_db_{edit_d_id}")
                                ed_v_tdrop_b = st.text_input("下硅膠墊垂落時間 (sec)", value=v_defs.get("下硅膠墊垂落時間 (sec)", ""), key=f"ed_v_tdb_{edit_d_id}")
                                c_v7, c_v8 = st.columns(2)
                                with c_v7: ed_v_pb = st.text_input("下加壓壓力 (kgf/cm²)", value=v_defs.get("下加壓壓力 (kgf/cm²)", ""), key=f"ed_v_pb_{edit_d_id}")
                                with c_v8: ed_v_tpb = st.text_input("下加壓時間 (sec)", value=v_defs.get("下加壓時間 (sec)", ""), key=f"ed_v_tpb_{edit_d_id}")
                        else:
                            pre_defs = unpack_params(row_dict.get('Pre_Lam', ''))
                            with st.expander("📍 預貼機參數", expanded=True):
                                c_p1, c_p2 = st.columns(2)
                                with c_p1: ed_pre_t = st.text_input("預貼溫度 (℃)", value=pre_defs.get("預貼溫度 (℃)", ""), key=f"ed_p_t_{edit_d_id}")
                                with c_p2: ed_pre_s = st.text_input("預貼速度 (m/min)", value=pre_defs.get("預貼速度 (m/min)", ""), key=f"ed_p_s_{edit_d_id}")
                                c_p3, c_p4 = st.columns(2)
                                with c_p3: ed_pre_p = st.text_input("預貼壓力 (MPa)", value=pre_defs.get("預貼壓力 (MPa)", ""), key=f"ed_p_p_{edit_d_id}")
                                with c_p4: ed_pre_m = st.text_input("前後留邊量 (前mm / 後mm)", value=pre_defs.get("前後留邊量", ""), key=f"ed_p_m_{edit_d_id}")
                                    
                            ed_l1_dict = render_lam_inputs("1st 壓模機", "el1", edit_d_id, unpack_params(row_dict.get('Lam_1st', '')))
                            ed_l2_dict = render_lam_inputs("2nd 壓模機", "el2", edit_d_id, unpack_params(row_dict.get('Lam_2nd', '')))
                            ed_l3_dict = render_lam3_inputs("3rd 壓模機", "el3", edit_d_id, unpack_params(row_dict.get('Lam_3rd', '')))
                        
                        st.write("---")
                        c_q1, c_q2 = st.columns(2)
                        with c_q1: ed_qty = st.text_input("壓合數量 (片/次)", value=str(row_dict.get('Qty', '')), key=f"ed_qty_{edit_d_id}")
                        
                        eval_opts = ["⚪ 尚未評估", "🟢 佳 (參數可參考)", "🟡 普通 (需微調)", "🔴 差 (不建議使用)"]
                        old_eval = str(row_dict.get('Self_Eval', ''))
                        e_idx = eval_opts.index(old_eval) if old_eval in eval_opts else 0
                        with c_q2: ed_eval = st.selectbox("內部自評結果", eval_opts, index=e_idx, key=f"ed_eval_{edit_d_id}")
                        
                        ed_remark = st.text_area("備註 (測試變動說明、具體異常)", value=str(row_dict.get('Remarks', '')), key=f"ed_rmk_{edit_d_id}")
                        ed_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)", value=str(row_dict.get('Feedback', '')), key=f"ed_fb_{edit_d_id}")
                        ed_upload = st.file_uploader("🖼️ 更新測試照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"ed_photo_{edit_d_id}")
                        
                        st.write("---")
                        edit_d_msg = st.empty()
                        if st.form_submit_button("💾 覆蓋更新實驗紀錄", key=f"btn_ed_{edit_d_id}"):
                            with st.spinner("打包與更新雲端資料庫中..."):
                                new_photo_url = upload_image(ed_upload, f"{edit_d_id}_edit.jpg") if ed_upload else str(row_dict.get('Photo_URL', ''))
                                final_ed_sub_t = ed_sub_t_other if ed_sub_t == "其他" else ed_sub_t
                                final_ed_film_m = ed_film_m_other if ed_film_m == "其他" else ed_film_m
                                
                                if is_v160:
                                    new_v_dict = {
                                        "加壓模式": ed_v_mode, "下真空時間 (sec)": ed_v_tv, "上溫度 (℃)": ed_v_tt, "下溫度 (℃)": ed_v_tb,
                                        "上硅膠墊垂落時間 (sec)": ed_v_tdrop_t, "上氣囊加壓壓力 (kgf/cm²)": ed_v_pt, "上氣囊加壓時間 (sec)": ed_v_tpt,
                                        "下加壓延遲時間 (sec)": ed_v_dly_b, "下硅膠墊垂落時間 (sec)": ed_v_tdrop_b,
                                        "下加壓壓力 (kgf/cm²)": ed_v_pb, "下加壓時間 (sec)": ed_v_tpb
                                    }
                                    final_pre, final_l1, final_l2, final_l3 = "無", pack_params(new_v_dict), "無", "無"
                                else:
                                    final_pre = pack_params({"預貼溫度 (℃)": ed_pre_t, "預貼壓力 (MPa)": ed_pre_p, "預貼速度 (m/min)": ed_pre_s, "前後留邊量": ed_pre_m})
                                    final_l1 = pack_params(ed_l1_dict)
                                    final_l2 = pack_params(ed_l2_dict)
                                    final_l3 = pack_params(ed_l3_dict)

                                new_d_row = [edit_d_id, ed_date.strftime("%Y-%m-%d"), ed_operator, ed_customer, ed_equip, final_ed_sub_t, ed_sub_size, final_ed_film_m, ed_film_model, final_pre, final_l1, final_l2, final_l3, ed_qty, ed_eval, ed_remark, ed_feedback, new_photo_url]
                                cell = sheet_demo.find(edit_d_id, in_column=1)
                                sheet_demo.update(values=[new_d_row], range_name=f"A{cell.row}:R{cell.row}")
                                st.cache_data.clear()
                                edit_d_msg.success(f"✅ 實驗單號 {edit_d_id} 更新成功！")
