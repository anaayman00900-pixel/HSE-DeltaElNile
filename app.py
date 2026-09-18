import streamlit as st
import pandas as pd
import numpy as np

# إعدادات الصفحة والتنسيق العام
st.set_page_config(page_title="نظام إدارة السلامة والصحة المهنية", page_icon="🛡️", layout="wide")

# تطبيق تنسيق CSS لتحسين المظهر ودعم اللغة العربية (RTL)
st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    div[data-testid="stMetricValue"] { font-size: 24px; font-weight: bold; color: #1E88E5; }
    .stDataFrame { border: 1px solid #e0e0e0; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("🛡️ لوحة تحكم السلامة والصحة المهنية - دلتا النيل")

excel_file = "Delta_Nile_HSE_Management_System_V2 (Recovered).xlsm"

def clean_dataframe(df):
    """دالة لتنظيف الجداول والتخلص من الخلايا الدمج والسطور الفارغة"""
    # إزالة الصفوف والأعمدة الفارغة تمامًا
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # البحث عن أول سطر يحتوي على عناوين حقيقية
    header_idx = 0
    for idx, row in df.iterrows():
        non_null_count = row.dropna().count()
        if non_null_count > 2:  # افترضنا أن سطر الهيدر يحتوي على أكثر من عنصرين
            header_idx = idx
            break
            
    # إعادة ضبط الهيدر
    if header_idx > 0:
        new_header = df.iloc[header_idx].astype(str).str.strip()
        df = df.iloc[header_idx + 1:].copy()
        df.columns = new_header

    # تنظيف أسماء الأعمدة من Unnamed وقيم None
    clean_cols = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if "Unnamed" in col_str or col_str == "None" or col_str == "nan":
            clean_cols.append(f"عمود_{i+1}")
        else:
            clean_cols.append(col_str)
    df.columns = clean_cols
    
    # إستبدال قيم NaN بالنصوص الفارغة للعرض
    df = df.replace({np.nan: "-", "None": "-", "nan": "-"})
    return df

@st.cache_data
def load_all_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheets_dict = {}
    for sheet_name in xls.sheet_names:
        try:
            raw_df = xls.parse(sheet_name)
            cleaned_df = clean_dataframe(raw_df)
            if not cleaned_df.empty:
                sheets_dict[sheet_name] = cleaned_df
        except Exception:
            continue
    return sheets_dict

try:
    data_sheets = load_all_data(excel_file)
    
    # شريط التحكم الجانبي
    st.sidebar.image("https://img.icons8.com/color/96/000000/worker-safety.png", width=80)
    st.sidebar.header("قائمة السجلات والشيتات")
    selected_sheet = st.sidebar.selectbox("اختر الشيت المطلوب:", list(data_sheets.keys()))
    
    df = data_sheets[selected_sheet]
    
    # كروت مؤشرات الأداء السريعة (KPIs)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("إجمالي السجلات (الصفوف)", len(df))
    with col2:
        st.metric("عدد البيانات (الأعمدة)", len(df.columns))
    with col3:
        st.metric("السجل المفتوح حالياً", selected_sheet)
        
    st.markdown("---")
    
    # شريط البحث والتصفية
    st.subheader(f"📊 عرض البيانات: {selected_sheet}")
    search_query = st.text_input("🔍 بحث سريع في هذا السجل:")
    
    if search_query:
        mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        filtered_df = df[mask]
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"حدث خطأ أثناء معالجة البيانات: {e}")
