import streamlit as st
import pandas as pd

# ฟังก์ชันสำหรับโหลดข้อมูลจากไฟล์ Excel
def load_excel(file):
    sheets = pd.ExcelFile(file).sheet_names  # ดึงรายชื่อชีททั้งหมด
    dataframes = {sheet: pd.read_excel(file, sheet_name=sheet) for sheet in sheets}  # โหลดข้อมูลจากทุกชีท
    return dataframes

# ฟังก์ชันสำหรับค้นหาแบบอ่อนตัว
def search_data(dataframes, query):
    result = {}
    for sheet_name, df in dataframes.items():
        # สร้าง DataFrame ใหม่ที่มีเฉพาะเซลล์ที่มีคำค้นหา
        filtered_df = df[df.applymap(lambda cell: query.lower() in str(cell).lower())]
        if not filtered_df.isnull().all().all():  # ตรวจสอบว่า DataFrame มีข้อมูลตรงกับคำค้นหาหรือไม่
            result[sheet_name] = filtered_df
    return result

# ส่วน UI ของ Streamlit
st.title("Excel Viewer and Fuzzy Search Tool")
st.subheader("Upload an Excel file to display all sheets and perform a fuzzy search.")

# อัพโหลดไฟล์ Excel
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # โหลดข้อมูลจากไฟล์
    dataframes = load_excel(uploaded_file)
    st.success("Excel file loaded successfully!")

    # แสดงข้อมูลในแต่ละชีท
    st.subheader("Data Preview:")
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
                # เนื่องจากจะได้เฉพาะเซลล์ที่แมตช์ ต้องแสดงผลพร้อม index
                st.dataframe(result_df.style.highlight_null(null_color='gray'))
        else:
            st.warning("No results found for your query.")
