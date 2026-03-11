import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
import requests
import base64
import plotly.express as px

# 設定台灣時區 (UTC+8)
tz_tw = timezone(timedelta(hours=8))

# --- 初始化狀態記憶體 ---
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# 📌 你的 Google Apps Script 專屬接收站網址
GAS_URL = "https://script.google.com/macros/s/AKfycbxEVcNlZjjFEmkQmH8Ft-P8mVTSQllsfFF0Khf4YE8lmuOvRQBU8lzocmFs04oMm6g5/exec"

# --- 🛠️ 輔助功能區 (打包參數與顯示過濾) ---
def pack_params(param_dict):
    lines = []
    for k, v in param_dict.items():
        if v and str(v).strip():
            lines.append(f"{k}：{v.strip()}")
    return "\n".join(lines) if lines else "無"

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

# 📌 1st / 2nd 壓模機專屬輸入 UI (加入了動態 Key 確保成功後清空)
def render_lam_inputs(stage_name, key_prefix, fk):
    with st.expander(f"📍 {stage_name} 參數"):
        c1, c2 = st.columns(2)
        with c1: t = st.text_input("溫度 (℃)", key=f"{key_prefix}_t_{fk}")
        with c2: t_v = st.text_input("抽真空時間 (sec)", key=f"{key_prefix}_tv_{fk}")
            
        c3, c4 = st.columns(2)
        with c3: p = st.text_input("加壓壓力 (kgf/cm²)", key=f"{key_prefix}_p_{fk}")
        with c4: t_p = st.text_input("加壓時間 (sec)", key=f"{key_prefix}_tp_{fk}")
            
        return {
            "溫度 (℃)": t, "抽真空時間 (sec)": t_v,
            "加壓壓力 (kgf/cm²)": p, "加壓時間 (sec)": t_p
        }

# 📌 3rd 壓模機 (伺服控制) 專屬輸入 UI (加入了動態 Key 確保成功後清空)
def render_lam3_inputs(stage_name, key_prefix, fk):
    with st.expander(f"📍 {stage_name} 參數 (伺服控制)"):
        mode = st.selectbox("控制模式", ["", "Position", "Press", "Fit"], key=f"{key_prefix}_mode_{fk}")
        st.write("---")
        c1, c2 = st.columns(2)
        with c1: t = st.text_input("溫度 (℃)", key=f"{key_prefix}_t_{fk}")
        with c2: t_v = st.text_input("抽真空時間 (sec)", key=f"{key_prefix}_tv_{fk}")
        c3, c4 = st.columns(2)
        with c3: thick = st.text_input("目前產品厚度 (mm)", key=f"{key_prefix}_thk_{fk}")
        st.markdown("###### 🎯 模式專屬參數 (請對應上方模式填寫，未填將自動隱藏)")
        c5, c6, c7 = st.columns(3)
        with c5: pos_v = st.text_input("【Position】厚度補償", key=f"{key_prefix}_pos_{fk}")
        with c6: press_v = st.text_input("【Press】加壓壓力", key=f"{key_prefix}_prs_{fk}")
        with c7: fit_v = st.text_input("【Fit】推進量", key=f"{key_prefix}_fit_{fk}")
        st.write("---")
        c8, c9 = st.columns(2)
        with c8: spd = st.text_input("加壓推速度 (mm/sec)", key=f"{key_prefix}_spd_{fk}")
        with c9: t_p = st.text_input("加壓時間 (sec)", key=f"{key_prefix}_tp_{fk}")
            
        return {
            "控制模式": mode, "溫度 (℃)": t, "抽真空時間 (sec)": t_v, "目前產品厚度 (mm)": thick, 
            "厚度補償 (Position)": pos_v, "加壓壓力 (Press)": press_v, "推進量 (Fit)": fit_v,
            "加壓推速度 (mm/sec)": spd, "加壓時間 (sec)": t_p
        }
# ---------------------------------------------

# 1. 取得金鑰並連線到 Google Sheets
@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_maint = gc.open("設備維修知識庫").worksheet("維修紀錄")
    sheet_demo = gc.open("設備維修知識庫").worksheet("實驗參數")
    return sheet_maint, sheet_demo

sheet_maint, sheet_demo = init_connection()

# 2. 透過秘密通道 (GAS) 上傳照片
def upload_image(image_file, file_name):
    base64_image = base64.b64encode(image_file.getvalue()).decode('utf-8')
    payload = {"fileName": file_name, "mimeType": image_file.type, "fileData": base64_image}
    return requests.post(GAS_URL, data=payload).text 

# 3. 讀取試算表資料
@st.cache_data(ttl=60)
def load_data(mode):
    if mode == "maint":
        data = sheet_maint.get_all_records()
    else:
        data = sheet_demo.get_all_records()
    df = pd.DataFrame(data)
    if not df.empty:
        df = df.iloc[::-1].reset_index(drop=True)
    return df

# --- 側邊欄：雙系統模式切換 ---
with st.sidebar:
    st.markdown("### 🎛️ 系統模式切換")
    app_mode = st.radio("選擇要使用的系統：", ["🔧 現場維修系統", "🧪 DEMO 實驗紀錄"], label_visibility="collapsed")
    st.write("---")
    st.markdown("### 📴 無塵室離線準備")
    if app_mode == "🔧 現場維修系統":
        df_maint = load_data("maint")
        if not df_maint.empty:
            st.download_button("📥 下載維修離線版 (CSV)", data=df_maint.to_csv(index=False).encode('utf-8-sig'), file_name=f"維修紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
    else:
        df_demo = load_data("demo")
        if not df_demo.empty:
            st.download_button("📥 下載實驗離線版 (CSV)", data=df_demo.to_csv(index=False).encode('utf-8-sig'), file_name=f"實驗紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
# -----------------------------

# --- 標題區塊 ---
col1, col2 = st.columns([1, 5])
with col1:
    try: st.image("logo.png", width=80) 
    except: st.title("⚙️") 
with col2:
    st.markdown(f"<h1 style='margin-top: -15px;'>{'設備維修知識庫' if app_mode == '🔧 現場維修系統' else 'DEMO 實驗資料庫'}</h1>", unsafe_allow_html=True)
# -----------------------------

# 共用 CSS 樣式
st.markdown("""
<style>
.glide-card { background-color: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 12px; border-left: 6px solid #FFA726; }
.glide-title { font-size: 16px; font-weight: 700; color: #333333; margin-bottom: 4px; }
.glide-subtitle { font-size: 14px; color: #555555; margin-bottom: 10px; line-height: 1.4; }
.glide-tag { background-color: #FFF3E0; color: #E65100; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; margin-bottom: 6px; }
.glide-solution { font-size: 13px; color: #D84315; background-color: #FBE9E7; padding: 8px; border-radius: 6px; margin-top: 6px; }
.glide-img { width: 100%; border-radius: 8px; margin-top: 10px; border: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 模式 A：現場維修系統
# ==========================================
if app_mode == "🔧 現場維修系統":
    tab1, tab2, tab3 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄", "📊 數據分析"])
    df = load_data("maint")

    with tab1:
        search_keyword = st.text_input("🔍 全域搜尋 (例如: 日期, 廠區, 問題 ... 等 關鍵字)")
        filtered_df = df.copy()

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
                        photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                        val_issue = str(row.get('Issue_Desc', '')).replace('\r', '').replace('\n', '<br>')
                        val_solution = str(row.get('Solution', '')).replace('\r', '').replace('\n', '<br>')
                            
                        st.markdown(f"""
<div class="glide-card">
<div class="glide-title">{row.get('Component', '')}</div>
<div class="glide-tag">📅 {row.get('Date', '')}</div>
<div class="glide-tag">🏢 {row.get('Customer', '')}</div>
<div class="glide-tag">⚙️ {row.get('Machine_Model', '')}</div>
<div class="glide-tag">👤 {row.get('Engineer', '')}</div>
<div class="glide-subtitle"><b>狀況：</b><br>{val_issue}</div>
<div class="glide-solution"><b>💡 解法：</b><br>{val_solution}</div>
{photo_html}
</div>
""", unsafe_allow_html=True)
        else:
            st.warning("目前試算表中沒有資料喔！")

    with tab2:
        fk = st.session_state.form_key
        with st.form(f"maint_form_{fk}"):
            st.subheader("📝 填寫維修紀錄")
            comp_options = ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"]
            
            input_date = st.date_input("日期", datetime.now(tz_tw).date(), key=f"m_date_{fk}")
            input_engineer = st.text_input("填單人員", key=f"m_eng_{fk}")
            input_customer = st.text_input("客戶與廠區", key=f"m_cust_{fk}")
            input_machine = st.selectbox("設備機型", ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"], index=None, key=f"m_mach_{fk}")
            input_component = st.selectbox("異常部件", comp_options, index=None, key=f"m_comp_{fk}")
            input_issue = st.text_area("問題描述", key=f"m_iss_{fk}")
            input_solution = st.text_area("解決方案", key=f"m_sol_{fk}")
            upload_file = st.file_uploader("🖼️ 附加現場照片", type=['jpg', 'png', 'jpeg'], key=f"m_photo_{fk}")
            
            if st.form_submit_button("送出維修紀錄", key=f"btn_m_{fk}"):
                if not all([input_engineer, input_customer, input_machine, input_component, input_issue, input_solution]):
                    # 資料不齊全時亮紅燈，但 form_key 不增加，資料完美保留！
                    st.error("⚠️ 請確認所有必填欄位都已填寫！")
                else:
                    with st.spinner("寫入中..."):
                        log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M")
                        photo_url = upload_image(upload_file, f"{log_id}.jpg") if upload_file else ""
                        sheet_maint.append_row([log_id, input_date.strftime("%Y-%m-%d"), input_engineer, input_customer, input_machine, input_component, input_issue, input_solution, photo_url])
                        st.cache_data.clear()
                        st.session_state.success_msg = f"✅ 成功寫入資料庫！單號：{log_id}"
                        # 成功送出後，增加 form_key，觸發重新載入，表單自動清空！
                        st.session_state.form_key += 1
                        st.rerun()

        if st.session_state.success_msg:
            st.success(st.session_state.success_msg)
            st.session_state.success_msg = ""

    with tab3:
        st.subheader("📈 維修數據統計看板")
        if not df.empty:
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("累積維修總件數", f"{len(df)} 件")
            with col2: st.metric("本月新增件數", f"{df[df['Date'].str.startswith(datetime.now(tz_tw).strftime('%Y-%m'), na=False)].shape[0]} 件")
            with col3: st.metric("涵蓋機型數量", f"{df['Machine_Model'].nunique()} 種")
                
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                machine_counts = df['Machine_Model'].value_counts().reset_index()
                machine_counts.columns = ['機型', '次數']
                fig_pie = px.pie(machine_counts, names='機型', values='次數', hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrBr[2:])
                fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0), dragmode=False)
                st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

            with col_chart2:
                comp_counts = df['Component'].value_counts().reset_index()
                comp_counts.columns = ['部件', '次數']
                fig_bar = px.bar(comp_counts, x='次數', y='部件', orientation='h', text='次數', color_discrete_sequence=['#FFA726'])
                fig_bar.update_traces(textposition='outside')
                fig_bar.update_layout(yaxis={'categoryorder':'total ascending', 'fixedrange': True}, xaxis={'fixedrange': True}, margin=dict(t=0, b=0, l=0, r=0), xaxis_title="報修次數", yaxis_title="", dragmode=False)
                st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info("目前系統中還沒有資料，等輸入幾筆維修紀錄後，這裡就會自動變出圖表囉！")

# ==========================================
# 模式 B：DEMO 實驗紀錄
# ==========================================
elif app_mode == "🧪 DEMO 實驗紀錄":
    tab_d1, tab_d2, tab_d3 = st.tabs(["🔍 參數查詢", "➕ 新增實驗紀錄 (NT+CVP)", "➕ 新增實驗紀錄 (V-160)"])
    df_d = load_data("demo")

    with tab_d1:
        search_kw_demo = st.text_input("🔍 全域搜尋 (例如: 膜材型號, 客戶, 機台)")
        filtered_demo = df_d.copy()

        if not filtered_demo.empty:
            if search_kw_demo:
                mask = pd.Series(False, index=filtered_demo.index)
                for col in filtered_demo.columns: mask = mask | filtered_demo[col].astype(str).str.contains(search_kw_demo, case=False, na=False)
                filtered_demo = filtered_demo[mask]
            
            st.caption(f"🔍 找到 {len(filtered_demo)} 筆實驗紀錄")
            
            for index, row in filtered_demo.iterrows():
                photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                
                equip_name = str(row.get('Equipment', ''))
                html_sub = format_params_html(row.get('Substrate_Info', ''))
                html_pre = format_params_html(row.get('Pre_Lam', ''))
                html_l1 = format_params_html(row.get('Lam_1st', ''))
                html_l2 = format_params_html(row.get('Lam_2nd', ''))
                html_l3 = format_params_html(row.get('Lam_3rd', ''))
                
                blocks = []
                if "V-160" in equip_name.upper() or "V160" in equip_name.upper():
                    if html_l1: blocks.append(f"<b>🔹 V-160 參數：</b><br>{html_l1}")
                else:
                    if html_pre: blocks.append(f"<b>🔹 預貼機參數：</b><br>{html_pre}")
                    if html_l1: blocks.append(f"<b>🔹 1st 壓模：</b><br>{html_l1}")
                    if html_l2: blocks.append(f"<b>🔹 2nd 壓模：</b><br>{html_l2}")
                    if html_l3: blocks.append(f"<b>🔹 3rd 壓模：</b><br>{html_l3}")
                
                params_block = ""
                if blocks:
                    inner_html = "<br><br>".join(blocks)
                    params_block = f"<div style='background-color:#F9F9F9; padding:10px; border-radius:8px; margin-bottom:10px; font-size:13px; color:#555;'>{inner_html}</div>"
                    
                sub_block = f"<div class='glide-subtitle'><b>基材/膜材：</b><br>{html_sub}</div>" if html_sub else ""

                st.markdown(f"""
<div class="glide-card">
<div class="glide-title">🧪 測試機台: {equip_name if equip_name else '未填寫'}</div>
<div class="glide-tag">📅 {row.get('Date', '')}</div>
<div class="glide-tag">🏢 {row.get('Customer', '')}</div>
<div class="glide-tag">👤 操作: {row.get('Operator', '')}</div>
<div class="glide-tag">📦 數量: {row.get('Qty', '')}</div>
{sub_block}
{params_block}
<div class="glide-subtitle"><b>📝 備註與異常：</b><br>{str(row.get('Remarks', '無')).replace('\n', '<br>')}</div>
<div class="glide-solution"><b>🗣️ 客戶反饋：</b><br>{str(row.get('Feedback', '無')).replace('\n', '<br>')}</div>
{photo_html}
</div>
""", unsafe_allow_html=True)
        else:
            st.info("目前還沒有 DEMO 實驗紀錄，趕快去新增一筆吧！")

    with tab_d2:
        fk = st.session_state.form_key
        with st.form(f"demo_form_{fk}"):
            st.subheader("🧪 填寫實驗紀錄 (NT+CVP)")
            c1, c2 = st.columns(2)
            with c1: input_d_date = st.date_input("測試日期", datetime.now(tz_tw).date(), key=f"d_date_{fk}")
            with c2: input_d_customer = st.text_input("客戶名稱", key=f"d_cust_{fk}")
            c4, c5 = st.columns(2)
            with c4: input_d_operator = st.text_input("操作人", key=f"d_oper_{fk}")
            with c5: input_d_equip = st.text_input("設備類型 (如: CVP-1600SP)", key=f"d_equip_{fk}")
            
            with st.expander("📍 基材資訊 (沒填寫的將自動隱藏)"):
                c_s1, c_s2 = st.columns(2)
                with c_s1: sub_t = st.text_input("板材類型", key=f"s_t_{fk}")
                with c_s2: sub_f = st.text_input("膜材 (供應商/型號/厚度)", key=f"s_f_{fk}")
                sub_d = st.text_input("基板尺寸 (大小/厚度)", key=f"s_d_{fk}")
            
            st.write("---")
            st.markdown("##### ⚙️ 各站機台參數設定 (沒填寫的欄位將自動隱藏)")
            
            with st.expander("📍 預貼機參數"):
                c_p1, c_p2 = st.columns(2)
                with c_p1: pre_t = st.text_input("預貼溫度 (℃)", key=f"p_t_{fk}")
                with c_p2: pre_s = st.text_input("預貼速度 (m/min)", key=f"p_s_{fk}")
                c_p3, c_p4 = st.columns(2)
                with c_p3: pre_p = st.text_input("預貼壓力 (MPa)", key=f"p_p_{fk}")
                with c_p4: pre_m = st.text_input("前後留邊量 (前mm / 後mm)", key=f"p_m_{fk}")
                    
            lam1_dict = render_lam_inputs("1st 壓模機", "l1", fk)
            lam2_dict = render_lam_inputs("2nd 壓模機", "l2", fk)
            lam3_dict = render_lam3_inputs("3rd 壓模機", "l3", fk)
                
            st.write("---")
            input_d_qty = st.text_input("壓合數量 (片/次)", key=f"d_qty_{fk}")
            input_d_remark = st.text_area("備註 (測試變動說明、具體異常)", key=f"d_rmk_{fk}")
            input_d_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)", key=f"d_fb_{fk}")
            upload_d_file = st.file_uploader("🖼️ 附加測試結果照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"d_photo_{fk}")
            
            if st.form_submit_button("送出實驗紀錄", key=f"btn_d_{fk}"):
                if not all([input_d_operator, input_d_customer, input_d_equip]):
                    st.error("⚠️ 請至少填寫操作人、客戶名稱與設備類型！")
                else:
                    with st.spinner("打包參數並寫入雲端中..."):
                        input_d_substrate = pack_params({"板材類型": sub_t, "膜材": sub_f, "基板尺寸": sub_d})
                        input_d_pre = pack_params({"預貼溫度 (℃)": pre_t, "預貼壓力 (MPa)": pre_p, "預貼速度 (m/min)": pre_s, "前後留邊量": pre_m})
                        input_d_1st = pack_params(lam1_dict)
                        input_d_2nd = pack_params(lam2_dict)
                        input_d_3rd = pack_params(lam3_dict)

                        log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        photo_url = upload_image(upload_d_file, f"{log_id}.jpg") if upload_d_file else ""
                        
                        new_demo_row = [
                            log_id, input_d_date.strftime("%Y-%m-%d"), input_d_operator, 
                            input_d_customer, input_d_equip, input_d_substrate, 
                            input_d_pre, input_d_1st, input_d_2nd, input_d_3rd, 
                            input_d_qty, input_d_remark, input_d_feedback, photo_url
                        ]
                        
                        sheet_demo.append_row(new_demo_row)
                        st.cache_data.clear()
                        st.session_state.success_msg = f"✅ 成功寫入 DEMO 紀錄！單號：{log_id}"
                        st.session_state.form_key += 1
                        st.rerun()

    with tab_d3:
        fk = st.session_state.form_key
        with st.form(f"v160_form_{fk}"):
            st.subheader("🧪 填寫實驗紀錄 (V-160)")
            
            c1, c2 = st.columns(2)
            with c1: input_v_date = st.date_input("測試日期", datetime.now(tz_tw).date(), key=f"v_date_{fk}")
            with c2: input_v_customer = st.text_input("客戶名稱", key=f"v_cust_{fk}")
            
            c4, c5 = st.columns(2)
            with c4: input_v_operator = st.text_input("操作人", key=f"v_oper_{fk}")
            with c5: input_v_equip = st.text_input("設備類型", value="V-160", disabled=True, key=f"v_equip_{fk}") 
            
            with st.expander("📍 基材資訊 (沒填寫的將自動隱藏)"):
                c_vs1, c_vs2 = st.columns(2)
                with c_vs1: v_sub_t = st.text_input("板材類型", key=f"v_s_t_{fk}")
                with c_vs2: v_sub_f = st.text_input("膜材 (供應商/型號/厚度)", key=f"v_s_f_{fk}")
                v_sub_d = st.text_input("基板尺寸 (大小/厚度)", key=f"v_s_d_{fk}")
            
            st.write("---")
            
            with st.expander("📍 V-160 參數"):
                c_v1, c_v2 = st.columns(2)
                with c_v1: v_mode = st.selectbox("加壓模式", ["", "上", "下", "上下"], key=f"v_mode_{fk}")
                with c_v2: v_tv = st.text_input("下真空時間 (sec)", key=f"v_tv_{fk}")
                
                c_v3, c_v4 = st.columns(2)
                with c_v3: v_tt = st.text_input("上溫度 (℃)", key=f"v_tt_{fk}")
                with c_v4: v_tb = st.text_input("下溫度 (℃)", key=f"v_tb_{fk}")
                
                st.write("---")
                v_tdrop_t = st.text_input("上硅膠墊垂落時間 (sec)", key=f"v_tdrop_t_{fk}")
                
                c_v5, c_v6 = st.columns(2)
                with c_v5: v_pt = st.text_input("上氣囊加壓壓力 (kgf/cm²)", key=f"v_pt_{fk}")
                with c_v6: v_tpt = st.text_input("上氣囊加壓時間 (sec)", key=f"v_tpt_{fk}")
                
                st.write("---")
                v_dly_b = st.text_input("下加壓延遲時間 (sec)", key=f"v_dly_b_{fk}")
                v_tdrop_b = st.text_input("下硅膠墊垂落時間 (sec)", key=f"v_tdrop_b_{fk}")
                
                c_v7, c_v8 = st.columns(2)
                with c_v7: v_pb = st.text_input("下加壓壓力 (kgf/cm²)", key=f"v_pb_{fk}")
                with c_v8: v_tpb = st.text_input("下加壓時間 (sec)", key=f"v_tpb_{fk}")
                
            st.write("---")
            input_v_qty = st.text_input("壓合數量 (片/次)", key=f"v_qty_{fk}")
            input_v_remark = st.text_area("備註 (測試變動說明、具體異常)", key=f"v_rmk_{fk}")
            input_v_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)", key=f"v_fb_{fk}")
            upload_v_file = st.file_uploader("🖼️ 附加測試結果照片 (選填)", type=['jpg', 'png', 'jpeg'], key=f"v_photo_{fk}")
            
            # 📌 V-160 的按鈕名稱統一為「送出實驗紀錄」
            if st.form_submit_button("送出實驗紀錄", key=f"btn_v_{fk}"):
                if not all([input_v_operator, input_v_customer]):
                    st.error("⚠️ 請至少填寫操作人與客戶名稱！")
                else:
                    with st.spinner("打包 V-160 參數並寫入雲端中..."):
                        input_v_substrate = pack_params({"板材類型": v_sub_t, "膜材": v_sub_f, "基板尺寸": v_sub_d})
                        
                        v160_dict = {
                            "加壓模式": v_mode,
                            "下真空時間 (sec)": v_tv,
                            "上溫度 (℃)": v_tt,
                            "下溫度 (℃)": v_tb,
                            "上硅膠墊垂落時間 (sec)": v_tdrop_t,
                            "上氣囊加壓壓力 (kgf/cm²)": v_pt,
                            "上氣囊加壓時間 (sec)": v_tpt,
                            "下加壓延遲時間 (sec)": v_dly_b,
                            "下硅膠墊垂落時間 (sec)": v_tdrop_b,
                            "下加壓壓力 (kgf/cm²)": v_pb,
                            "下加壓時間 (sec)": v_tpb
                        }
                        input_v_params = pack_params(v160_dict)

                        log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        photo_url = upload_image(upload_v_file, f"{log_id}.jpg") if upload_v_file else ""
                        
                        new_v160_row = [
                            log_id, input_v_date.strftime("%Y-%m-%d"), input_v_operator, 
                            input_v_customer, input_v_equip, input_v_substrate, 
                            "無", input_v_params, "無", "無", 
                            input_v_qty, input_v_remark, input_v_feedback, photo_url
                        ]
                        
                        sheet_demo.append_row(new_v160_row)
                        st.cache_data.clear()
                        st.session_state.success_msg = f"✅ 成功寫入 V-160 紀錄！單號：{log_id}"
                        st.session_state.form_key += 1
                        st.rerun()

    # 📌 共用顯示成功訊息
    if st.session_state.success_msg:
        st.success(st.session_state.success_msg)
        st.session_state.success_msg = ""
