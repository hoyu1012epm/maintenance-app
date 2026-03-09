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

# 1. 取得金鑰並連線到 Google Sheets
@st.cache_resource 
def init_connection():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet = gc.open("設備維修知識庫").sheet1
    return sheet

sheet = init_connection()

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

# 3. 讀取試算表資料
@st.cache_data(ttl=60)
def load_data():
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()
if not df.empty:
    df = df.iloc[::-1].reset_index(drop=True)

# --- 側邊欄：離線急救包下載 ---
with st.sidebar:
    st.markdown("### 📴 無塵室離線準備")
    st.caption("進無塵室前，請先下載最新紀錄至手機")
    
    if not df.empty:
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載完整離線版 (CSV)",
            data=csv,
            file_name=f"維修知識庫_離線版_{datetime.now(tz_tw).strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
# -----------------------------

# --- 標題區塊：完美對齊版 ---
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image("logo.png", width=80) 
    except:
        st.title("🔧") 
with col2:
    st.markdown("<h1 style='margin-top: -15px;'>設備維修知識庫</h1>", unsafe_allow_html=True)
# -----------------------------

tab1, tab2, tab3 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄", "📊 數據分析"])

# ==========================================
# 分頁 1：查詢紀錄 (恢復為直觀的一目了然版)
# ==========================================
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

        # 🎨 恢復為原本的直觀卡片 CSS
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
                        
                    # 📌 恢復完全展開的 HTML 結構 (靠左對齊防跑版)
                    st.markdown(f"""
<div class="glide-card">
<div class="glide-title">{row['Component']}</div>
<div class="glide-tag">📅 {row['Date']}</div>
<div class="glide-tag">🏢 {row['Customer']}</div>
<div class="glide-tag">⚙️ {row['Machine_Model']}</div>
<div class="glide-tag">👤 {row['Engineer']}</div>
<div class="glide-subtitle"><b>狀況：</b>{row['Issue_Desc']}</div>
<div class="glide-solution"><b>💡 解法：</b>{row['Solution']}</div>
{photo_html}
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
        
        st.write("---")
        upload_file = st.file_uploader("🖼️ 附加現場照片 (選填)", type=['jpg', 'png', 'jpeg'])
