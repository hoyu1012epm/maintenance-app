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

# 1. 取得金鑰並連線到 Google Sheets (同時載入兩張表)
@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    # 📌 連線到兩張不同的工作表
    sheet_maint = gc.open("設備維修知識庫").worksheet("維修紀錄")
    sheet_demo = gc.open("設備維修知識庫").worksheet("實驗參數")
    return sheet_maint, sheet_demo

sheet_maint, sheet_demo = init_connection()

# 2. 透過秘密通道 (GAS) 上傳照片
def upload_image(image_file, file_name):
    base64_image = base64.b64encode(image_file.getvalue()).decode('utf-8')
    payload = {
        "fileName": file_name,
        "mimeType": image_file.type,
        "fileData": base64_image
    }
    response = requests.post(GAS_URL, data=payload)
    return response.text 

# 3. 讀取試算表資料 (依據模式讀取不同的表)
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
    app_mode = st.radio(
        "選擇要使用的系統：",
        ["🔧 現場維修系統", "🧪 DEMO 實驗紀錄"],
        label_visibility="collapsed"
    )
    
    st.write("---")
    st.markdown("### 📴 無塵室離線準備")
    
    if app_mode == "🔧 現場維修系統":
        df_maint = load_data("maint")
        if not df_maint.empty:
            csv_maint = df_maint.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載維修離線版 (CSV)", data=csv_maint, file_name=f"維修紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
    else:
        df_demo = load_data("demo")
        if not df_demo.empty:
            csv_demo = df_demo.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載實驗離線版 (CSV)", data=csv_demo, file_name=f"實驗紀錄_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv", mime="text/csv")
# -----------------------------

# --- 標題區塊 ---
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image("logo.png", width=80) 
    except:
        st.title("⚙️") 
with col2:
    if app_mode == "🔧 現場維修系統":
        st.markdown("<h1 style='margin-top: -15px;'>設備維修知識庫</h1>", unsafe_allow_html=True)
    else:
        st.markdown("<h1 style='margin-top: -15px;'>DEMO 實驗資料庫</h1>", unsafe_allow_html=True)
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

    # 📌 完美修復：把原本的分類功能加回來了！
    with tab1:
        search_keyword = st.text_input("🔍 全域搜尋 (例如: 日期, 廠區, 問題 ... 等 關鍵字)")
        filtered_df = df.copy()

        if not filtered_df.empty:
            try:
                filtered_df['YearMonth'] = pd.to_datetime(filtered_df['Date']).dt.strftime('%y/%m')
            except Exception:
                filtered_df['YearMonth'] = "未知時間"

            if search_keyword:
                mask = pd.Series(False, index=filtered_df.index)
                for col in filtered_df.columns:
                    mask = mask | filtered_df[col].astype(str).str.contains(search_keyword, case=False, na=False)
                filtered_df = filtered_df[mask]

            st.write("---")
            group_by_option = st.radio(
                "🗂️ 選擇展開分類方式：",
                ["依建立年月", "依客戶與廠區", "依設備機型", "依設備部件"], 
                horizontal=True
            )

            group_col_map = {
                "依建立年月": "YearMonth", 
                "依客戶與廠區": "Customer",
                "依設備機型": "Machine_Model",
                "依設備部件": "Component" 
            }
            group_col = group_col_map[group_by_option]
            
            st.caption(f"🔍 找到 {len(filtered_df)} 筆相關紀錄")

            custom_component_order = [
                "預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", 
                "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", 
                "PLC", "真空/氣壓系統", "溫控系統", "其他"
            ]

            unique_groups = filtered_df[group_col].unique().tolist()

            if group_col == "Component":
                unique_groups.sort(key=lambda x: custom_component_order.index(x) if x in custom_component_order else 999)
            elif group_col == "YearMonth":
                unique_groups.sort(reverse=True) 
            else:
                unique_groups.sort(key=lambda x: str(x))

            for group_name in unique_groups:
                display_name = group_name if str(group_name).strip() != "" else "未分類/未填寫"
                group_data = filtered_df[filtered_df[group_col] == group_name]
                
                with st.expander(f"📁 {display_name} (共 {len(group_data)} 筆)"):
                    for index, row in group_data.iterrows():
                        photo_html = ""
                        if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http"):
                            photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">'
                            
                        st.markdown(f"""
<div class="glide-card">
<div class="glide-title">{row.get('Component', '')}</div>
<div class="glide-tag">📅 {row.get('Date', '')}</div>
<div class="glide-tag">🏢 {row.get('Customer', '')}</div>
<div class="glide-tag">⚙️ {row.get('Machine_Model', '')}</div>
<div class="glide-tag">👤 {row.get('Engineer', '')}</div>
<div class="glide-subtitle"><b>狀況：</b>{row.get('Issue_Desc', '')}</div>
<div class="glide-solution"><b>💡 解法：</b>{row.get('Solution', '')}</div>
{photo_html}
</div>
""", unsafe_allow_html=True)

        else:
            st.warning("目前試算表中沒有資料喔！")

    with tab2:
        with st.form(f"maint_form_{st.session_state.form_key}"):
            st.subheader("📝 填寫維修紀錄")
            comp_options = ["預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"]
            
            input_date = st.date_input("日期", datetime.now(tz_tw).date())
            input_engineer = st.text_input("填單人員")
            input_customer = st.text_input("客戶與廠區")
            input_machine = st.selectbox("設備機型", ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"], index=None)
            input_component = st.selectbox("異常部件", comp_options, index=None)
            input_issue = st.text_area("問題描述")
            input_solution = st.text_area("解決方案")
            upload_file = st.file_uploader("🖼️ 附加現場照片", type=['jpg', 'png', 'jpeg'])
            
            if st.form_submit_button("送出維修紀錄"):
                if not all([input_engineer, input_customer, input_machine, input_component, input_issue, input_solution]):
                    st.error("⚠️ 請確認所有必填欄位都已填寫！")
                else:
                    with st.spinner("寫入中..."):
                        log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M%S")
                        photo_url = upload_image(upload_file, f"{log_id}.jpg") if upload_file else ""
                        sheet_maint.append_row([log_id, input_date.strftime("%Y-%m-%d"), input_engineer, input_customer, input_machine, input_component, input_issue, input_solution, photo_url])
                        st.cache_data.clear()
                        st.session_state.success_msg = f"✅ 成功寫入資料庫！單號：{log_id}"
                        st.session_state.form_key += 1
                        st.rerun()

    with tab3:
        st.subheader("📈 維修數據統計看板")
        if not df.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("累積維修總件數", f"{len(df)} 件")
            with col2:
                current_month = datetime.now(tz_tw).strftime("%Y-%m")
                this_month_count = df[df['Date'].str.startswith(current_month, na=False)].shape[0]
                st.metric("本月新增件數", f"{this_month_count} 件")
            with col3:
                unique_machines = df['Machine_Model'].nunique()
                st.metric("涵蓋機型數量", f"{unique_machines} 種")
                
            st.write("---")
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.markdown("##### ⚙️ 各機型報修佔比")
                machine_counts = df['Machine_Model'].value_counts().reset_index()
                machine_counts.columns = ['機型', '次數']
                fig_pie = px.pie(machine_counts, names='機型', values='次數', hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrBr[2:])
                fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0), dragmode=False)
                st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

            with col_chart2:
                st.markdown("##### 🔧 異常部件排行榜")
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
    tab_d1, tab_d2 = st.tabs(["🔍 參數查詢", "➕ 新增實驗紀錄"])
    df_d = load_data("demo")

    with tab_d1:
        search_kw_demo = st.text_input("🔍 全域搜尋 (例如: 配方號碼, 膜材型號, 客戶)")
        filtered_demo = df_d.copy()

        if not filtered_demo.empty:
            if search_kw_demo:
                mask = pd.Series(False, index=filtered_demo.index)
                for col in filtered_demo.columns:
                    mask = mask | filtered_demo[col].astype(str).str.contains(search_kw_demo, case=False, na=False)
                filtered_demo = filtered_demo[mask]
            
            st.caption(f"🔍 找到 {len(filtered_demo)} 筆實驗紀錄")
            
            for index, row in filtered_demo.iterrows():
                photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">' if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http") else ""
                
                # DEMO 專用卡片排版
                st.markdown(f"""
                <div class="glide-card">
                <div class="glide-title">🧪 配方: {row.get('Recipe_NO', '未命名')} | 機台: {row.get('Equipment', '')}</div>
                <div class="glide-tag">📅 {row.get('Date', '')}</div>
                <div class="glide-tag">🏢 {row.get('Customer', '')}</div>
                <div class="glide-tag">👤 操作: {row.get('Operator', '')}</div>
                <div class="glide-tag">📦 數量: {row.get('Qty', '')}</div>
                <div class="glide-subtitle"><b>基材/膜材：</b>{row.get('Substrate_Info', '')}</div>
                
                <div style="background-color:#F9F9F9; padding:10px; border-radius:8px; margin-bottom:10px; font-size:13px; color:#555;">
                    <b>🔹 預貼機參數：</b> {row.get('Pre_Lam', '無')}<br>
                    <b>🔹 1st 壓模：</b> {row.get('Lam_1st', '無')}<br>
                    <b>🔹 2nd 壓模：</b> {row.get('Lam_2nd', '無')}<br>
                    <b>🔹 3rd 壓模：</b> {row.get('Lam_3rd', '無')}
                </div>
                
                <div class="glide-subtitle"><b>📝 備註與異常：</b>{row.get('Remarks', '無')}</div>
                <div class="glide-solution"><b>🗣️ 客戶反饋：</b>{row.get('Feedback', '無')}</div>
                {photo_html}
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("目前還沒有 DEMO 實驗紀錄，趕快去新增一筆吧！")

    with tab_d2:
        with st.form(f"demo_form_{st.session_state.form_key}"):
            st.subheader("🧪 填寫 DEMO 機測試紀錄")
            
            c1, c2, c3 = st.columns(3)
            with c1: input_d_date = st.date_input("測試日期", datetime.now(tz_tw).date())
            with c2: input_d_operator = st.text_input("操作人")
            with c3: input_d_applicant = st.text_input("申請人")
            
            c4, c5 = st.columns(2)
            with c4: input_d_customer = st.text_input("客戶名稱")
            with c5: input_d_equip = st.text_input("設備類型 (如: CVP-1600SP)")
            
            input_d_recipe = st.text_input("配方 NO. (Recipe)")
            input_d_substrate = st.text_area("基材資訊 (板材類型 / 膜材型號厚度 / 基板大小)")
            
            st.write("---")
            st.markdown("##### ⚙️ 各站機台參數設定")
            with st.expander("📍 預貼機參數"):
                input_d_pre = st.text_area("預貼機設定 (溫度/壓力/速度/空調等)", key="d_pre")
            with st.expander("📍 1st 壓模機參數"):
                input_d_1st = st.text_area("1st 壓模設定", key="d_1st")
            with st.expander("📍 2nd 壓模機參數"):
                input_d_2nd = st.text_area("2nd 壓模設定", key="d_2nd")
            with st.expander("📍 3rd 壓模機參數"):
                input_d_3rd = st.text_area("3rd 壓模設定", key="d_3rd")
                
            st.write("---")
            input_d_qty = st.text_input("壓合數量 (片/次)")
            input_d_remark = st.text_area("備註 (測試變動說明、具體異常)")
            input_d_feedback = st.text_area("客戶反饋 (Pass/Fail/改善點)")
            
            upload_d_file = st.file_uploader("🖼️ 附加測試結果照片 (選填)", type=['jpg', 'png', 'jpeg'], key="d_photo")
            
            if st.form_submit_button("送出實驗紀錄"):
                if not all([input_d_operator, input_d_customer, input_d_equip]):
                    st.error("⚠️ 請至少填寫操作人、客戶名稱與設備類型！")
                else:
                    with st.spinner("寫入實驗數據中..."):
                        log_id = datetime.now(tz_tw).strftime("DEMO-%y%m%d-%H%M")
                        photo_url = upload_image(upload_d_file, f"{log_id}.jpg") if upload_d_file else ""
                        
                        new_demo_row = [
                            log_id, input_d_date.strftime("%Y-%m-%d"), input_d_operator, input_d_applicant, 
                            input_d_customer, input_d_equip, input_d_recipe, input_d_substrate, 
                            input_d_pre, input_d_1st, input_d_2nd, input_d_3rd, 
                            input_d_qty, input_d_remark, input_d_feedback, photo_url
                        ]
                        
                        sheet_demo.append_row(new_demo_row)
                        st.cache_data.clear()
                        st.session_state.success_msg = f"✅ 成功寫入 DEMO 紀錄！單號：{log_id}"
                        st.session_state.form_key += 1
                        st.rerun()

# --- 顯示成功訊息 ---
if st.session_state.success_msg:
    st.success(st.session_state.success_msg)
    st.session_state.success_msg = ""
