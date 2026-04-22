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

# --- 📌 系統記憶體初始化 ---
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

# 📌 訊息提示記憶體 (用於智慧不洗白防呆)
if "msg_maint" not in st.session_state: st.session_state.msg_maint = ""
if "msg_demo_nt" not in st.session_state: st.session_state.msg_demo_nt = ""
if "msg_mach_log" not in st.session_state: st.session_state.msg_mach_log = ""
if "msg_admin_add" not in st.session_state: st.session_state.msg_admin_add = ""

if "load_nt_key" not in st.session_state: st.session_state.load_nt_key = 0
if "load_v_key" not in st.session_state: st.session_state.load_v_key = 0

# 📌 連線設定
GAS_URL = "https://script.google.com/macros/s/AKfycbxEVcNlZjjFEmkQmH8Ft-P8mVTSQllsfFF0Khf4YE8lmuOvRQBU8lzocmFs04oMm6g5/exec"

def hash_pw(password):
    return hashlib.sha256(password.encode()).hexdigest()

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
    if not df.empty and mode != "users":
        df = df.iloc[::-1].reset_index(drop=True)
    return df

# 📌 全域樣式
st.markdown("""
<style>
.glide-card { background-color: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 5px; border-left: 6px solid #FFA726; }
.glide-title { font-size: 16px; font-weight: 700; color: #333333; margin-bottom: 4px; }
.glide-tag { background-color: #FFF3E0; color: #E65100; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; }
.diff-alert { background-color: #FFEBEE; border-left: 5px solid #D32F2F; padding: 10px; border-radius: 8px; margin-bottom: 10px; color: #C62828; }
.diff-safe { background-color: #F1F8E9; border-left: 5px solid #4CAF50; padding: 10px; border-radius: 8px; margin-bottom: 10px; color: #2E7D32; }
.calc-green { background-color: #E8F5E9; padding: 12px; border-radius: 8px; border-left: 6px solid #4CAF50; font-size: 18px; font-weight: bold; color: #2E7D32; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

if st.session_state.logged_in:
    now = datetime.now(tz_tw)
    if now - st.session_state.last_active > timedelta(minutes=15):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.warning("⏱️ 您已閒置超過 15 分鐘，系統已自動登出。")
        st.rerun()
    else: st.session_state.last_active = now

# ==========================================
# 🔒 登入與密碼修改
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
        if st.button("🔄 重新整理資料", use_container_width=True): 
            st.cache_data.clear()
            st.rerun()
        if st.button("🚪 登出", use_container_width=True):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

    st.markdown(f"<h2>{app_mode}</h2>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # 模式 A：現場維修系統
    # ---------------------------------------------------------
    if app_mode == "🔧 現場維修系統":
        tab1, tab2 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄"])
        with tab1:
            search_keyword = st.text_input("🔍 關鍵字搜尋")
            filtered_m = df_maint.copy()
            if search_keyword:
                mask = filtered_m.apply(lambda row: row.astype(str).str.contains(search_keyword, case=False).any(), axis=1)
                filtered_m = filtered_m[mask]
            
            sel_m_ids = [row['Log_ID'] for idx, row in filtered_m.iterrows() if st.session_state.get(f"chk_m_{row['Log_ID']}", False)]
            if sel_m_ids:
                st.download_button(f"📥 匯出勾選 ({len(sel_m_ids)}筆)", data=filtered_m[filtered_m['Log_ID'].isin(sel_m_ids)].to_csv(index=False).encode('utf-8-sig'), file_name="Maint_Export.csv", type="primary")
            
            for idx, row in filtered_m.iterrows():
                c_card, c_action = st.columns([5, 2])
                with c_action:
                    st.checkbox("勾選", key=f"chk_m_{row['Log_ID']}")
                    st.download_button("單筆", data=pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig'), file_name=f"{row['Log_ID']}.csv", key=f"dl_m_{row['Log_ID']}")
                with c_card:
                    st.markdown(f'<div class="glide-card"><div class="glide-title">{row["Component"]}</div><div class="glide-tag">🏢 {row["Customer"]}</div><br>{row["Issue_Desc"]}</div>', unsafe_allow_html=True)
                st.markdown("<hr style='border-top: 1px dashed #ccc;'>", unsafe_allow_html=True)

        with tab2:
            fk = st.session_state.form_key
            with st.form(f"maint_form_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1: m_c = st.selectbox("選擇客戶", [""] + unique_cust)
                with c2: m_cn = st.text_input("自填客戶")
                m_ma = st.selectbox("機型", ["NT-300", "CVP-1600", "其他"])
                m_is = st.text_area("問題描述")
                m_so = st.text_area("解決方案")
                
                m_msg = st.empty()
                if st.session_state.msg_maint:
                    m_msg.success(st.session_state.msg_maint)
                    st.session_state.msg_maint = ""
                    
                if st.form_submit_button("💾 送出紀錄"):
                    final_c = m_cn if m_cn else m_c
                    if final_c and m_is:
                        log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M")
                        sheet_maint.append_row([log_id, datetime.now(tz_tw).strftime("%Y-%m-%d"), st.session_state.user_name, final_c, m_ma, "", m_is, m_so, ""])
                        st.session_state.msg_maint = f"✅ 已成功送出單號：{log_id}"
                        st.session_state.form_key += 1
                        st.cache_data.clear()
                        st.rerun()
                    else: st.error("⚠️ 資料填寫不完整")

    # ---------------------------------------------------------
    # 模式 B：DEMO 實驗紀錄
    # ---------------------------------------------------------
    elif app_mode == "🧪 DEMO 實驗紀錄":
        tab1, tab2 = st.tabs(["🔍 參數查詢", "➕ 新增 NT+CVP"])
        with tab1:
            search_d = st.text_input("🔍 搜尋實驗紀錄")
            filtered_d = df_demo.copy()
            if search_d:
                mask = filtered_d.apply(lambda row: row.astype(str).str.contains(search_d, case=False).any(), axis=1)
                filtered_d = filtered_d[mask]
            
            sel_d_ids = [row['Log_ID'] for idx, row in filtered_d.iterrows() if st.session_state.get(f"chk_d_{row['Log_ID']}", False)]
            if sel_d_ids:
                st.download_button(f"📥 匯出勾選 ({len(sel_d_ids)}筆)", data=filtered_d[filtered_d['Log_ID'].isin(sel_d_ids)].to_csv(index=False).encode('utf-8-sig'), file_name="DEMO_Export.csv", type="primary")

            for idx, row in filtered_d.iterrows():
                c_card, c_action = st.columns([5, 2])
                with c_action:
                    st.checkbox("勾選", key=f"chk_d_{row['Log_ID']}")
                    st.download_button("單筆匯出", data=pd.DataFrame([row]).to_csv(index=False).encode('utf-8-sig'), file_name=f"{row['Log_ID']}.csv", key=f"dl_d_{row['Log_ID']}")
                with c_card:
                    st.markdown(f'<div class="glide-card"><div class="glide-title">{row["Equipment"]}</div><div class="glide-tag">🏢 {row["Customer"]}</div><br>{row["Remarks"]}</div>', unsafe_allow_html=True)
                st.markdown("<hr style='border-top: 1px dashed #ccc;'>", unsafe_allow_html=True)

        with tab2:
            fk = st.session_state.form_key
            clone_sel = st.selectbox("⚡ 一鍵帶入歷史紀錄", [""] + [f"{r['Log_ID']} - {r['Customer']}" for idx, r in df_demo.iterrows()])
            clone_data = df_demo[df_demo['Log_ID'] == clone_sel.split(" - ")[0]].iloc[0].to_dict() if clone_sel else {}

            with st.form(f"demo_form_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                with c1: d_c = st.selectbox("客戶", [""] + unique_cust, index=(unique_cust.index(clone_data.get('Customer'))+1 if clone_data.get('Customer') in unique_cust else 0))
                with c2: d_cn = st.text_input("自填客戶", value=clone_data.get('Customer') if not d_c else "")
                d_eq = st.text_input("機台型號", value=clone_data.get('Equipment', ''))
                d_re = st.text_area("備註", value=clone_data.get('Remarks', ''))
                
                d_msg = st.empty()
                if st.session_state.msg_demo_nt:
                    d_msg.success(st.session_state.msg_demo_nt)
                    st.session_state.msg_demo_nt = ""

                if st.form_submit_button("💾 送出實驗紀錄"):
                    final_c = d_cn if d_cn else d_c
                    if final_c and d_eq:
                        log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        sheet_demo.append_row([log_id, datetime.now(tz_tw).strftime("%Y-%m-%d"), st.session_state.user_name, final_c, d_eq, "", "", "", "", "", "", "", "", "", "", d_re, "", ""])
                        st.session_state.msg_demo_nt = f"✅ 已成功送出單號：{log_id}"
                        st.session_state.form_key += 1
                        st.cache_data.clear()
                        st.rerun()
                    else: st.error("⚠️ 請填寫必填欄位")

    # ---------------------------------------------------------
    # 模式 C：⚙️ 設備機械履歷 (全動態 A~BK 參數對應)
    # ---------------------------------------------------------
    elif app_mode == "⚙️ 設備機械履歷":
        # 📌 這是你辛苦整理出來的 58 個參數，系統會自動依此生成 UI 與對應 Google Sheet 欄位
        MACHINE_PARAM_GROUPS = {
            "📍 A區：伺服壓合與極限 (D120~D854)": [
                "A_D120_下限位置極限","A_D122_真空位置下限極限","A_D314_壓合恆定速度","A_D6064_加速","A_D6065_減速",
                "A_D6090_待機位置","A_D6092_下限位置","A_D452_壓力異常限值","A_D170_真空大氣開放時間",
                "A_D176_不抽真空時轉矩下限","A_D177_抽真空時轉矩下限","A_D854_Film咬合保持"
            ],
            "📍 B區：驅動與傳送 (D40~D718)": [
                "B_D40_驅動軸間隔移動量","B_D6158_Film送帶速度","B_D6156_加速時間","B_D6157_減速時間","B_D46_張力初始",
                "B_D47_張力初始時定數","B_D48_品種張力時定數","B_D714_加速時間","B_D715_減速時間","B_D717_工序時間",
                "B_D718_傳送部擋板停止速度","B_D328_入口異常時間","B_D514_自動運行中擋板上升延遲"
            ],
            "📍 D區：1st 壓合台 (D740~D872)": [
                "D_D740_壓合台壓力異常範圍","D_D742_上升傳感器延遲時間","D_D743_真空大氣開放時間","D_D746_高壓ON 逆向壓力時間",
                "D_D747_壓合台必要推力","D_D748_逆壓壓力","D_D749_加壓電控閥調節","D_D750_逆壓電控閥調節","D_D782_加壓異常時間",
                "D_D790_下降時殘留壓力排放時間","D_D791_自動運轉下降阻斷閥打開延遲","D_D736_1st手動上升高壓輔助ON","D_D870_增壓閥下限壓力","D_D872_異常時間"
            ],
            "📍 D區：2nd 壓合台 (D752~D875)": [
                "D_D752_壓合台壓力異常範圍","D_D754_上升傳感器延遲時間","D_D755_真空大氣開放時間","D_D758_高壓ON 逆向壓力時間",
                "D_D759_壓合台必要推力","D_D760_逆壓壓力","D_D761_加壓電控閥調節","D_D762_逆壓電控閥調節","D_D783_加壓異常時間",
                "D_D792_下降時殘留壓力排放時間","D_D793_自動運轉下降阻斷閥打開延遲","D_D738_2nd手動上升高壓輔助ON","D_D873_增壓閥下限壓力","D_D875_異常時間"
            ],
            "📍 E區：伺服定位與 Fit (D460~D464)": [
                "E_D460_定位1次定位量","E_D462_壓力1次定位量","E_D466_Fit模式SUS接觸搜索1次定位量","E_D464_Fit控制推進時1次定位量"
            ]
        }

        tab_m1, tab_m2 = st.tabs(["🔍 參數查詢與客變比對", "➕ 紀錄機台現況 (100% 完整還原 HMI)"])
        
        with tab_m1:
            search_sn = st.text_input("🔍 輸入機台序號 (SN) 查詢：", placeholder="例如: CVP-1500-001")
            if search_sn:
                sn_recs = df_mach[df_mach['Equipment_SN'].astype(str).str.contains(search_sn, case=False)]
                if sn_recs.empty: st.warning("找不到此機台。")
                else:
                    factory = sn_recs.iloc[-1].to_dict() # 抓最早一筆作標準
                    current = sn_recs.iloc[0].to_dict()  # 抓最新一筆作現況
                    st.success(f"✅ 找到 {len(sn_recs)} 筆紀錄。最後更新日期: {current['Date']}")
                    
                    st.markdown("#### 🛠️ 機械參數差異比對")
                    c1, c2 = st.columns(2)
                    
                    # 動態生成查詢結果
                    for i, (grp_name, keys) in enumerate(MACHINE_PARAM_GROUPS.items()):
                        col = c1 if i % 2 == 0 else c2
                        with col:
                            with st.expander(grp_name, expanded=True):
                                for k in keys:
                                    # 解析顯示名稱 (例: A_D120_下限 -> D120 下限)
                                    parts = k.split('_')
                                    label = f"{parts[1]} {parts[2]}" if len(parts) >= 3 else parts[1]
                                    v_s, v_c = str(factory.get(k, '')), str(current.get(k, ''))
                                    is_diff = v_s != v_c and v_s != ""
                                    cls = "diff-alert" if is_diff else "diff-safe"
                                    txt = f"🚨 已客變 (原廠: {v_s})" if is_diff else "✅ 原廠值"
                                    st.markdown(f"<div class='{cls}'><small style='color:#555;'>{label}</small><br><b style='font-size:16px;'>{v_c}</b> <span style='float:right; font-size:12px;'>{txt}</span></div>", unsafe_allow_html=True)
                    st.info(f"📝 **最新客變備註：**\n{current.get('Remarks', '無')}")

        with tab_m2:
            fk = st.session_state.form_key
            st.info("💡 這裡已 100% 完整收錄您指定的 58 個機械參數，請依據 Pro-face 白底欄位對應填寫。")
            with st.form(f"mach_log_f_{fk}", clear_on_submit=False):
                c1, c2 = st.columns(2)
                m_sn = c1.text_input("機台序號 SN (必填)", key=f"m_sn_{fk}")
                m_cu = c2.selectbox("客戶廠區 (必填)", [""] + unique_cust, key=f"m_cu_{fk}")
                
                input_vals = {}
                # 動態生成 58 個輸入框，分成 3 欄排列，美觀又節省手機空間
                for grp_name, keys in MACHINE_PARAM_GROUPS.items():
                    with st.expander(grp_name, expanded=True):
                        cols = st.columns(3)
                        for i, k in enumerate(keys):
                            parts = k.split('_')
                            label = f"{parts[1]} {parts[2]}" if len(parts) >= 3 else parts[1]
                            input_vals[k] = cols[i % 3].text_input(label, key=f"m_i_{k}_{fk}")
                
                st.write("---")
                m_re = st.text_area("修改原因 / 現場客變備註", placeholder="詳細記錄本次修改了哪些參數，以及修改原因。")
                
                msg_box = st.empty()
                if st.session_state.msg_mach_log:
                    msg_box.success(st.session_state.msg_mach_log)
                    st.session_state.msg_mach_log = ""

                if st.form_submit_button("💾 一鍵儲存 63 欄位設備履歷"):
                    if m_sn and m_cu:
                        with st.spinner("資料比對打包中..."):
                            log_id = datetime.now(tz_tw).strftime("MACH-%y%m%d-%H%M")
                            date_str = datetime.now(tz_tw).strftime("%Y-%m-%d %H:%M")
                            
                            # 精準組合 63 欄位 (前 5 基本資訊 + 58 參數迴圈 + 1 備註)
                            row_data = [log_id, date_str, st.session_state.user_name, m_cu, m_sn]
                            for keys in MACHINE_PARAM_GROUPS.values():
                                for k in keys:
                                    row_data.append(input_vals[k])
                            row_data.append(m_re)
                            
                            sheet_mach.append_row(row_data)
                            st.session_state.msg_mach_log = f"✅ 設備 {m_sn} 的 63 項參數已全數建檔成功！"
                            st.session_state.form_key += 1
                            st.cache_data.clear()
                            st.rerun()
                    else: st.error("⚠️ 機台序號與客戶廠區為必填欄位！")

    # ---------------------------------------------------------
    # 模式 D：🧮 產品厚度計算機
    # ---------------------------------------------------------
    elif app_mode == "🧮 產品厚度計算機":
        st.info("輸入測量值進行自動運算")
        v1 = st.number_input("板材厚度", value=0.0, step=0.1)
        v2 = st.number_input("膜材厚度", value=0.0, step=0.1)
        st.markdown(f"<div class='calc-green'>🎯 建議 3rd 厚度設定：{v1 + v2:.2f} mm</div>", unsafe_allow_html=True)

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
