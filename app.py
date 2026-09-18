import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="HSE Management System", page_icon="📊", layout="wide")
st.title("📊 لوحة تحكم نظام السلامة والصحة المهنية")

excel_file = "Delta_Nile_HSE_Management_System_V2 (Recovered).xlsm"

@st.cache_data
def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheets_data = {}
    for sheet in xls.sheet_names:
        try:
            # قراءة الشيت بشكل مرن دون إجبار هيدر معين لتفادي أخطاء الشيتات القصيرة
            df = xls.parse(sheet)
            # تنظيف الأعمدة والصفوف الفارغة تماماً
            df = df.dropna(how='all').dropna(axis=1, how='all')
            sheets_data[sheet] = df
        except Exception as e:
            continue
    return sheets_data

try:
    data_sheets = load_data(excel_file)
    st.sidebar.success("تم تحميل البيانات بنجاح! ✅")
    
    st.sidebar.header("القائمة الرئيسية")
    selected_sheet = st.sidebar.selectbox("اختر الشيت للعرض:", list(data_sheets.keys()))
    
    df = data_sheets[selected_sheet]
    
    col1, col2, col3 = st.columns(3)
    col1.metric("إجمالي السجلات", len(df))
    col2.metric("عدد الأعمدة", len(df.columns))
    col3.metric("الشيت الحالي", selected_sheet)
    
    st.markdown("---")
    st.subheader(f"📋 جدول البيانات: {selected_sheet}")
    
    search_term = st.text_input("🔍 بحث داخل الشيت:")
    if search_term:
        filtered_df = df[df.astype(str).apply(lambda row: row.str.contains(search_term, case=False).any(), axis=1)]
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error(f"خطأ في تحميل الملف: {e}")
