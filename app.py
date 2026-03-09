import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials

# 1. 取得金鑰並連線到 Google Sheets
@st.cache_resource 
def init_connection():
    # 從 Streamlit 秘密金庫讀取剛剛貼上的 JSON
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

client = init_connection()

# 2. 讀取試算表資料
@st.cache_data(ttl=60) # 每60秒自動去試算表抓一次最新資料
def load_data():
    # 注意：這裡的 "設備維修知識庫" 必須跟你 Google 試算表的檔名一模一樣！
    sheet = client.open("設備維修知識庫").sheet1
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()

# 3. 介面設計與搜尋邏輯
st.title("🔧 設備維修知識庫 (雲端連線版)")

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

    st.subheader("📋 搜尋結果")
    st.dataframe(filtered_df, use_container_width=True)
else:
    st.warning("目前試算表中沒有資料，或是欄位名稱讀取失敗喔！")
