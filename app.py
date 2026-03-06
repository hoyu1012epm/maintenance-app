import streamlit as st
import pandas as pd

# 1. 建立測試資料庫 (這裡先用簡單的表格代替 SQLite，方便你馬上測試)
data = {
    "Log_ID": ["REP-001", "REP-002", "REP-003"],
    "Date": ["2026-02-13", "2026-02-26", "2026-03-05"],
    "Engineer": ["何宇", "Yamaguchi", "何宇"],
    "Machine_Model": ["CVP-1600SP", "CVP-1600SP", "預貼機"],
    "Component": ["壓模機-1st", "控制介面 (HMI)", "預貼機-投入"],
    "Issue_Desc": ["進行 ABF 壓膜時，邊緣產生不規則氣泡", "Pro-face 介面當機，無法連線", "感測器未觸發，導致卡料"],
    "Solution": ["調整真空度參數，並重新校正", "重新安裝 Pro-face 驅動程式", "清潔光電感測器"]
}
df = pd.DataFrame(data)

# 2. 設計網頁介面
st.title("🔧 設備維修知識庫")

# 建立兩個搜尋條件：下拉選單 與 關鍵字輸入
col1, col2 = st.columns(2)
with col1:
    # 抓取所有不重複的部件作為選單
    component_list = ["全部"] + list(df["Component"].unique())
    selected_comp = st.selectbox("請選擇設備部件", component_list)

with col2:
    search_keyword = st.text_input("輸入關鍵字 (例如: 氣泡, Pro-face)")

# 3. 執行搜尋與篩選邏輯
filtered_df = df.copy()

if selected_comp != "全部":
    filtered_df = filtered_df[filtered_df["Component"] == selected_comp]

if search_keyword:
    # 只要問題描述或解決方案包含關鍵字，就篩選出來
    mask = filtered_df["Issue_Desc"].str.contains(search_keyword, case=False) | \
           filtered_df["Solution"].str.contains(search_keyword, case=False)
    filtered_df = filtered_df[mask]

# 4. 顯示結果
st.subheader("📋 搜尋結果")
st.dataframe(filtered_df, use_container_width=True)
