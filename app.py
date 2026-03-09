import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone

# 設定台灣時區 (UTC+8)
tz_tw = timezone(timedelta(hours=8))

# --- 初始化表單狀態記憶體 ---
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# 1. 取得金鑰並連線到 Google Sheets
@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

client = init_connection()
sheet = client.open("設備維修知識庫").sheet1

# 2. 讀取試算表資料
@st.cache_data(ttl=60)
def load_data():
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()

# 將資料表上下顛倒，確保最新寫入的紀錄會排在最上面
if not df.empty:
    df = df.iloc[::-1].reset_index(drop=True)

# --- 標題區塊：調整欄位比例讓 LOGO 變大 ---
# 把比例從 [1, 6] 放大成 [1.5, 4]，給左邊的 LOGO 更多空間
col1, col2 = st.columns([1.5, 4]) 

with col1:
    try:
        # 拿掉指定的 width 數字，改用 use_container_width=True 讓圖片自動撐滿房間
        st.image("logo.png", use_container_width=True) 
    except:
        st.title("🔧") 

with col2:
    st.title("設備維修知識庫")
# -----------------------------

# 建立兩個分頁 (Tabs)
tab1, tab2 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄"])

# ==========================================
# 分頁 1：查詢紀錄 
# ==========================================
with tab1:
    search_keyword = st.text_input("🔍 全域搜尋 (例如: 日期, 廠區, 問題 ... 等 關鍵字)")

    filtered_df = df.copy()

    if not filtered_df.empty:
        # 📌 新增：動態將 Date 欄位 (2026-03-09) 轉換成年月格式 (26/03) 供分類使用
        try:
            filtered_df['YearMonth'] = pd.to_datetime(filtered_df['Date']).dt.strftime('%y/%m')
        except Exception:
            filtered_df['YearMonth'] = "未知時間"

        # 執行全欄位關鍵字搜尋
        if search_keyword:
            mask = pd.Series(False, index=filtered_df.index)
            for col in filtered_df.columns:
                mask = mask | filtered_df[col].astype(str).str.contains(search_keyword, case=False, na=False)
            filtered_df = filtered_df[mask]

        st.write("---")
        # 📌 修正：加入「依建立年月」選項
        group_by_option = st.radio(
            "🗂️ 選擇展開分類方式：",
            ["依建立年月", "依客戶與廠區", "依設備機型", "依設備部件"], 
            horizontal=True
        )

        group_col_map = {
            "依建立年月": "YearMonth", # 對應到剛剛動態產生的欄位
            "依客戶與廠區": "Customer",
            "依設備機型": "Machine_Model",
            "依設備部件": "Component" 
        }
        group_col = group_col_map[group_by_option]

        # 注入 Glide 卡片 CSS
        st.markdown("""
        <style>
        .glide-card {
            background-color: #ffffff;
            padding: 16px;
            border-radius: 12px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            margin-bottom: 12px;
            border-left: 6px solid #00c4b4;
        }
        .glide-title { font-size: 16px; font-weight: 700; color: #1f2937; margin-bottom: 4px; }
        .glide-subtitle { font-size: 14px; color: #4b5563; margin-bottom: 10px; line-height: 1.4; }
        .glide-tag {
            background-color: #f3f4f6; color: #374151; padding: 3px 8px;
            border-radius: 12px; font-size: 11px; display: inline-block;
            margin-right: 4px; margin-bottom: 6px;
        }
        .glide-solution {
            font-size: 13px; color: #065f46; background-color: #d1fae5;
            padding: 8px; border-radius: 6px; margin-top: 6px;
        }
        </style>
        """, unsafe_allow_html=True)

        st.caption(f"🔍 找到 {len(filtered_df)} 筆相關紀錄")

        custom_component_order = [
            "預貼機-投入", "預貼機-排出", "壓模機-卷出", "壓模機-1st", 
            "壓模機-2nd", "壓模機-3rd", "壓模機-卷收", "控制介面 (HMI)", 
            "PLC", "真空/氣壓系統", "溫控系統", "其他"
        ]

        unique_groups = filtered_df[group_col].unique().tolist()

        # 📌 修正：新增年月排序邏輯 (越新的月份排在越上面)
        if group_col == "Component":
            unique_groups.sort(key=lambda x: custom_component_order.index(x) if x in custom_component_order else 999)
        elif group_col == "YearMonth":
            unique_groups.sort(reverse=True) # 降冪排序，讓 26/03 排在 26/02 前面
        else:
            unique_groups.sort(key=lambda x: str(x))

        for group_name in unique_groups:
            display_name = group_name if str(group_name).strip() != "" else "未分類/未填寫"
            group_data = filtered_df[filtered_df[group_col] == group_name]
            
            with st.expander(f"📁 {display_name} (共 {len(group_data)} 筆)"):
                for index, row in group_data.iterrows():
                    st.markdown(f"""
                    <div class="glide-card">
                        <div class="glide-title">{row['Component']}</div>
                        <div class="glide-tag">📅 {row['Date']}</div>
                        <div class="glide-tag">🏢 {row['Customer']}</div>
                        <div class="glide-tag">⚙️ {row['Machine_Model']}</div>
                        <div class="glide-tag">👤 {row['Engineer']}</div>
                        <div class="glide-subtitle"><b>狀況：</b>{row['Issue_Desc']}</div>
                        <div class="glide-solution"><b>💡 解法：</b>{row['Solution']}</div>
                    </div>
                    """, unsafe_allow_html=True)

    else:
        st.warning("目前試算表中沒有資料喔！")

# ==========================================
# 分頁 2：新增紀錄
# ==========================================
with tab2:
    with st.form(f"add_record_form_{st.session_state.form_key}"):
        st.subheader("📝 填寫現場維修紀錄")
        
        comp_options = [
            "預貼機-投入", "預貼機-排出", 
            "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收",
            "控制介面 (HMI)", "PLC", "真空/氣壓系統", "溫控系統", "其他"
        ]
        
        input_date = st.date_input("日期", datetime.now(tz_tw).date())
        input_engineer = st.text_input("填單人員", placeholder="請輸入姓名")
        input_customer = st.text_input("客戶與廠區 (例如: 佰鼎 路竹)")
        input_machine = st.selectbox("設備機型", ["NT-300", "NT-400", "CVP-600", "CVP-1600", "CVP-1500", "其他"], index=None, placeholder="請選擇機型...")
        input_component = st.selectbox("發生異常的部件", comp_options, index=None, placeholder="請選擇部件...")
        input_issue = st.text_area("問題描述 (現象、錯誤代碼等)")
        input_solution = st.text_area("解決方案 (參數調整、更換零件等)")
        
        submitted = st.form_submit_button("送出紀錄至雲端")
        
        if submitted:
            missing_fields = []
            if not input_engineer: missing_fields.append("【填單人員】")
            if not input_customer: missing_fields.append("【客戶與廠區】")
            if not input_machine: missing_fields.append("【設備機型】")
            if not input_component: missing_fields.append("【發生異常的部件】")
            if not input_issue: missing_fields.append("【問題描述】")
            if not input_solution: missing_fields.append("【解決方案】")
            
            if missing_fields:
                st.error(f"⚠️ 提交失敗！請補充以下未填寫的欄位：{', '.join(missing_fields)}")
            else:
                log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M%S")
                date_str = input_date.strftime("%Y-%m-%d")
                
                new_row = [
                    log_id, date_str, input_engineer, 
                    input_customer, input_machine, input_component, 
                    input_issue, input_solution
                ]
                
                try:
                    sheet.append_row(new_row)
                    st.cache_data.clear() 
                    
                    st.session_state.success_msg = f"✅ 成功寫入資料庫！單號：{log_id}"
                    st.session_state.form_key += 1
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"寫入失敗，請檢查連線狀態：{e}")
                    
    if st.session_state.success_msg:
        st.success(st.session_state.success_msg)
        st.session_state.success_msg = ""





