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
# 分頁 1：查詢紀錄 (原有的功能)
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

        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.warning("目前試算表中沒有資料喔！")

# ==========================================
# 分頁 2：新增紀錄 (全新寫入功能)
# ==========================================
with tab2:
    # 建立表單，送出後自動清空欄位
    with st.form("add_record_form", clear_on_submit=True):
        st.subheader("📝 填寫現場維修紀錄")
        
        # 部件下拉選單清單
        comp_options = [
            "預貼機-投入", "預貼機-排出", 
            "壓模機-卷出", "壓模機-1st", "壓模機-2nd", "壓模機-3rd", "壓模機-卷收",
            "控制介面 (HMI)", "真空/氣壓系統", "溫控系統", "其他"
        ]
        
        # 表單輸入欄位配置
        input_date = st.date_input("日期", datetime.today())
        input_engineer = st.text_input("工程師", value="何宇")
        input_customer = st.text_input("客戶與廠區 (例如: A客戶 竹科廠)")
        input_machine = st.selectbox("設備機型", ["CVP-1600SP", "其他"])
        input_component = st.selectbox("發生異常的部件", comp_options)
        input_issue = st.text_area("問題描述 (現象、錯誤代碼等)")
        input_solution = st.text_area("解決方案 (參數調整、更換零件等)")
        
        # 建立送出按鈕
        submitted = st.form_submit_button("送出紀錄至雲端")
        
        # 當按下送出按鈕後的動作
        if submitted:
            # 簡單防呆：檢查重要欄位是否有填寫
            if not input_customer or not input_issue or not input_solution:
                st.error("⚠️ 請確認『客戶』、『問題描述』與『解決方案』都已填寫喔！")
            else:
                # 自動產生一組 Log_ID (例如: REP-20260309-1145)
                log_id = datetime.now().strftime("REP-%Y%m%d-%H%M")
                date_str = input_date.strftime("%Y-%m-%d")
                
                # 將收集到的資料打包成一個陣列 (順序必須與 Google 試算表欄位完全一致！)
                new_row = [
                    log_id, date_str, input_engineer, 
                    input_customer, input_machine, input_component, 
                    input_issue, input_solution
                ]
                
                try:
                    # 指令：將這筆資料寫入 Google 試算表的最下方
                    sheet.append_row(new_row)
                    st.success(f"✅ 成功寫入資料庫！單號：{log_id}")
                    
                    # 非常重要：清空 Streamlit 的快取，這樣切換回查詢頁面時才會抓到剛剛新增的資料
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f"寫入失敗，請檢查連線狀態：{e}")
