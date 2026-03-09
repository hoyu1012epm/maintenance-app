import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

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

# 將 sheet 設為全域變數，方便後續寫入資料
client = init_connection()
sheet = client.open("設備維修知識庫").sheet1

# 2. 讀取試算表資料
@st.cache_data(ttl=60)
def load_data():
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()

st.title("🔧 設備維修知識庫")

# 建立兩個分頁 (Tabs)
tab1, tab2 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄"])

# ==========================================
# 分頁 1：查詢紀錄 (導入 Glide 卡片美化風格)
# ==========================================
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        if not df.empty and "Component" in df.columns:
            component_list = ["全部"] + list(df["Component"].unique())
        else:
            component_list = ["全部"]
        selected_comp = st.selectbox("請選擇設備部件", component_list)

    with col2:
        search_keyword = st.text_input("輸入關鍵字 (例如: 氣泡, Pro-face)")

    filtered_df = df.copy()

    if not filtered_df.empty:
        if selected_comp != "全部":
            filtered_df = filtered_df[filtered_df["Component"] == selected_comp]

        if search_keyword:
            mask = filtered_df["Issue_Desc"].astype(str).str.contains(search_keyword, case=False, na=False) | \
                   filtered_df["Solution"].astype(str).str.contains(search_keyword, case=False, na=False)
            filtered_df = filtered_df[mask]

        # 注入自訂的 CSS 樣式 (打造 Glide 卡片風格)
        st.markdown("""
        <style>
        .glide-card {
            background-color: #ffffff;
            padding: 16px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05), 0 1px 3px rgba(0,0,0,0.1);
            margin-bottom: 16px;
            border-left: 6px solid #00c4b4; /* 左側的跳色裝飾線 */
        }
        .glide-title {
            font-size: 18px;
            font-weight: 700;
            color: #1f2937;
            margin-bottom: 4px;
        }
        .glide-subtitle {
            font-size: 15px;
            color: #4b5563;
            margin-bottom: 12px;
            line-height: 1.4;
        }
        .glide-tag {
            background-color: #f3f4f6;
            color: #374151;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 12px;
            display: inline-block;
            margin-right: 6px;
            margin-bottom: 8px;
        }
        .glide-solution {
            font-size: 14px;
            color: #065f46;
            background-color: #d1fae5;
            padding: 10px;
            border-radius: 8px;
            margin-top: 8px;
        }
        </style>
        """, unsafe_allow_html=True)

        # 顯示搜尋結果數量
        st.caption(f"🔍 找到 {len(filtered_df)} 筆相關紀錄")

        # 使用迴圈，把每一筆資料畫成一張獨立的卡片
        for index, row in filtered_df.iterrows():
            st.markdown(f"""
            <div class="glide-card">
                <div class="glide-title">{row['Component']}</div>
                <div class="glide-tag">📅 {row['Date']}</div>
                <div class="glide-tag">🏢 {row['Customer']}</div>
                <div class="glide-tag">⚙️ {row['Machine_Model']}</div>
                <div class="glide-subtitle"><b>狀況：</b>{row['Issue_Desc']}</div>
                <div class="glide-solution"><b>💡 解法：</b>{row['Solution']}</div>
            </div>
            """, unsafe_allow_html=True)

    else:
        st.warning("目前試算表中沒有資料喔！")

# ==========================================
# 分頁 2：新增紀錄 (寫入 Google 試算表)
# ==========================================
with tab2:
    with st.form("add_record_form", clear_on_submit=True):
        st.subheader("📝 填寫現場維修紀錄")
        
        comp_options = [
            "預貼機-投入", "預貼機-排出", 
            "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收",
            "控制介面 (HMI)", "真空/氣壓系統", "溫控系統", "其他"
        ]
        
        input_date = st.date_input("日期", datetime.today())
        input_engineer = st.text_input("工程師", value="何宇")
        input_customer = st.text_input("客戶與廠區 (例如: A客戶 竹科廠)")
        input_machine = st.selectbox("設備機型", ["CVP-1600SP", "其他"])
        input_component = st.selectbox("發生異常的部件", comp_options)
        input_issue = st.text_area("問題描述 (現象、錯誤代碼等)")
        input_solution = st.text_area("解決方案 (參數調整、更換零件等)")
        
        submitted = st.form_submit_button("送出紀錄至雲端")
        
        if submitted:
            if not input_customer or not input_issue or not input_solution:
                st.error("⚠️ 請確認『客戶』、『問題描述』與『解決方案』都已填寫喔！")
            else:
                log_id = datetime.now().strftime("REP-%Y%m%d-%H%M")
                date_str = input_date.strftime("%Y-%m-%d")
                
                new_row = [
                    log_id, date_str, input_engineer, 
                    input_customer, input_machine, input_component, 
                    input_issue, input_solution
                ]
                
                try:
                    sheet.append_row(new_row)
                    st.success(f"✅ 成功寫入資料庫！單號：{log_id}")
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"寫入失敗，請檢查連線狀態：{e}")
