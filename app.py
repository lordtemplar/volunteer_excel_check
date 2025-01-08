import streamlit as st
import pandas as pd

# ฟังก์ชันสำหรับโหลดข้อมูลจากไฟล์ Excel
def load_excel(file):
    sheets = pd.ExcelFile(file).sheet_names  # ดึงรายชื่อชีททั้งหมด
    dataframes = {}
    for sheet in sheets:
        df = pd.read_excel(file, sheet_name=sheet)
        # ลบข้อมูลในแถวที่ 0-3
        df = df.iloc[3:].reset_index(drop=True)
        dataframes[sheet] = df
    return dataframes

# ฟังก์ชันสำหรับค้นหาแบบอ่อนตัว
def search_data(dataframes, query):
    result = {}
    for sheet_name, df in dataframes.items():
        # ค้นหาแถวที่มีคำค้นหาในคอลัมน์ใดก็ได้
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False, na=False).any(), axis=1)]
        if not filtered_df.empty:  # ตรวจสอบว่า DataFrame มีข้อมูลตรงกับคำค้นหาหรือไม่
            result[sheet_name] = filtered_df
    return result

# ส่วน UI ของ Streamlit
st.title("Excel Viewer and Row-based Search Tool")
st.subheader("Upload an Excel file to display all sheets and search for rows with matching data.")

# อัพโหลดไฟล์ Excel
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # โหลดข้อมูลจากไฟล์
    dataframes = load_excel(uploaded_file)
    st.success("Excel file loaded successfully!")

    # แสดงข้อมูลในแต่ละชีท
    st.subheader("Data Preview (Rows 0-3 removed):")
    for sheet_name, df in dataframes.items():
        st.write(f"**Sheet: {sheet_name}**")
        st.dataframe(df)

    # ฟังก์ชันค้นหา
    query = st.text_input("Search for data (case insensitive):", "")
    if query:
        search_results = search_data(dataframes, query)
        if search_results:
            st.subheader("Search Results:")
            for sheet_name, result_df in search_results.items():
                st.write(f"**Sheet: {sheet_name}**")
                # แสดงเฉพาะแถวที่ค้นหาเจอ
                st.dataframe(result_df)
        else:
            st.warning("No results found for your query.")
