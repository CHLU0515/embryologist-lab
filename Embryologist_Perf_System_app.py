
import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO

# 預設資料儲存區（暫存在 Session State）
if 'records' not in st.session_state:
    st.session_state.records = []

st.title("🧬 胚胎師人員成效輸入系統 v0.1")

# 📝 輸入表單
with st.form("entry_form"):
    st.subheader("輸入每日操作數據")
    col1, col2 = st.columns(2)
    with col1:
        op_date = st.date_input("操作日期", value=date.today())
        operator = st.selectbox("胚胎師代號", ["00009", "00010", "00021", "00042", "00068"])
        retrievals = st.number_input("取卵台數", min_value=0)
        injection = st.number_input("顯微注射顆數 (ICSI)", min_value=0)
        fertilized = st.number_input("成功受精數", min_value=0)
    with col2:
        blastocysts = st.number_input("囊胚數 (Day 5)", min_value=0)
        frozen = st.number_input("冷凍胚數", min_value=0)
        sperm_proc = st.number_input("精蟲處理台數", min_value=0)
        denudation = st.number_input("剃蛋台數", min_value=0)
        media_prep = st.number_input("製備培養液台數", min_value=0)

    submitted = st.form_submit_button("📥 儲存紀錄")
    if submitted:
        record = {
            "日期": op_date,
            "代號": operator,
            "取卵台數": retrievals,
            "ICSI": injection,
            "受精數": fertilized,
            "受精率": round(fertilized / injection, 2) if injection else 0,
            "囊胚數": blastocysts,
            "囊胚形成率": round(blastocysts / fertilized, 2) if fertilized else 0,
            "冷凍胚數": frozen,
            "精蟲處理": sperm_proc,
            "剃蛋": denudation,
            "培養液製備": media_prep
        }
        st.session_state.records.append(record)
        st.success("已新增一筆資料 ✅")

# 📋 顯示現有資料與簡易統計
if st.session_state.records:
    df = pd.DataFrame(st.session_state.records)
    st.subheader("📊 紀錄總覽")
    st.dataframe(df)

    st.subheader("📈 成效平均統計")
    avg_data = df.groupby("代號")[["受精率", "囊胚形成率"]].mean().round(2)
    st.dataframe(avg_data)

    # 📤 匯出 Excel
    st.subheader("⬇️ 匯出報表")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='每日紀錄', index=False)
        avg_data.to_excel(writer, sheet_name='平均統計')
    buffer.seek(0)
    st.download_button(
        label="📤 下載 Excel 報表",
        data=buffer,
        file_name="胚胎師成效報表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )