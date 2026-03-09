import streamlit as st
import pandas as pd
import json
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import datetime, timedelta, timezone
import io

# 設定台灣時區 (UTC+8)
tz_tw = timezone(timedelta(hours=8))

# --- 初始化狀態記憶體 ---
if "form_key" not in st.session_state:
    st.session_state.form_key = 0
if "success_msg" not in st.session_state:
    st.session_state.success_msg = ""

# 📌 這是你的 Google Drive 維修照片庫 ID
FOLDER_ID = "1ayZ2qhsu5DMe9JRxkP0AO_QHesalpIdT"

# 1. 取得金鑰並連線到 Google Sheets & Google Drive
@st.cache_resource 
def init_connections():
    creds_dict = json.loads(st.secrets["gcp_credentials"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    
    # 連線到 Google Sheets
    gc = gspread.authorize(creds)
    sheet = gc.open("設備維修知識庫").sheet1
    
    # 連線到 Google Drive
    drive_service = build('drive', 'v3', credentials=creds)
    
    return sheet, drive_service

sheet, drive_service = init_connections()

# 2. 上傳照片至 Google Drive 的專屬函式
def upload_image(image_file, file_name):
    file_metadata = {
        'name': file_name,
        'parents': [FOLDER_ID]
    }
    media = MediaIoBaseUpload(io.BytesIO(image_file.getvalue()), 
                              mimetype=image_file.type, 
                              resumable=True)
    file = drive_service.files().create(body=file_metadata, 
                                        media_body=media, 
                                        fields='id').execute()
    file_id = file.get('id')
    # 回傳可以直接在網頁上顯示的圖片網址
    return f"https://drive.google.com/thumbnail?id={file_id}&sz=w1000"

# 3. 讀取試算表資料
@st.cache_data(ttl=60)
def load_data():
    data = sheet.get_all_records()
    return pd.DataFrame(data)

df = load_data()
if not df.empty:
    df = df.iloc[::-1].reset_index(drop=True)

# --- 標題區塊 ---
col1, col2 = st.columns([1, 6])
with col1:
    try:
        st.image("logo.png", width=100) 
    except:
        st.title("🔧") 
with col2:
    st.title("設備維修知識庫")
# -----------------------------

tab1, tab2 = st.tabs(["🔍 查詢紀錄", "➕ 新增紀錄"])

# ==========================================
# 分頁 1：查詢紀錄
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

        st.markdown("""
        <style>
        .glide-card { background-color: #ffffff; padding: 16px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 12px; border-left: 6px solid #D5896F; }
        .glide-title { font-size: 16px; font-weight: 700; color: #1f2937; margin-bottom: 4px; }
        .glide-subtitle { font-size: 14px; color: #5C4A44; margin-bottom: 10px; line-height: 1.4; }
        .glide-tag { background-color: #F4EBE6; color: #6B5B56; padding: 3px 8px; border-radius: 12px; font-size: 11px; display: inline-block; margin-right: 4px; margin-bottom: 6px; }
        .glide-solution { font-size: 13px; color: #8C4B31; background-color: #FAEBE6; padding: 8px; border-radius: 6px; margin-top: 6px; }
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
                    # 判斷有無照片網址，如果有就顯示圖片
                    photo_html = ""
                    if "Photo_URL" in row and str(row["Photo_URL"]).startswith("http"):
                        photo_html = f'<img src="{row["Photo_URL"]}" class="glide-img">'
                        
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
        
        # 📌 新增：純粹的照片上傳區塊 (沒有相機功能)
        st.write("---")
        upload_file = st.file_uploader("🖼️ 附加現場照片 (選填)", type=['jpg', 'png', 'jpeg'])
        
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
                try:
                    # 顯示轉圈圈動畫，讓使用者知道正在處理上傳
                    with st.spinner("正在安全寫入資料與上傳照片..."):
                        log_id = datetime.now(tz_tw).strftime("REP-%y%m%d-%H%M%S")
                        date_str = input_date.strftime("%Y-%m-%d")
                        
                        # 如果有上傳照片，就執行上傳並取得網址；沒有就留空
                        photo_url = ""
                        if upload_file is not None:
                            # 檔名自動取為 單號.副檔名
                            photo_url = upload_image(upload_file, f"{log_id}.jpg")
                        
                        # 📌 注意：多了一個 photo_url 寫入 Google Sheets
                        new_row = [
                            log_id, date_str, input_engineer, 
                            input_customer, input_machine, input_component, 
                            input_issue, input_solution, photo_url
                        ]
                        
                        sheet.append_row(new_row)
                        st.cache_data.clear() 
                        
                        st.session_state.success_msg = f"✅ 成功寫入資料庫！單號：{log_id}"
                        st.session_state.form_key += 1
                        st.rerun()
                        
                except Exception as e:
                    st.error(f"寫入失敗，請檢查連線或權限設定：{e}")
                    
    if st.session_state.success_msg:
        st.success(st.session_state.success_msg)
        st.session_state.success_msg = ""
