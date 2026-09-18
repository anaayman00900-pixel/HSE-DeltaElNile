import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# 1. إعدادات الصفحة والتصميم العادي
st.set_page_config(
    page_title="نظام إدارة السلامة والصحة المهنية - دلتا النيل",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. تحسين المظهر مع دعم اللغة العربية وتصميم متطور (CSS Custom Styling)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
        direction: rtl;
        text-align: right;
    }
    
    /* خلفية التطبيق */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* تصميم الكروت الإحصائية */
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        text-align: center;
        margin-bottom: 10px;
    }
    .metric-card h3 { margin: 0; font-size: 16px; opacity: 0.9; color: #e0e0e0; }
    .metric-card h2 { margin: 10px 0 0 0; font-size: 28px; font-weight: bold; }
    
    /* تصميم الجداول */
    .stDataFrame {
        background: white;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    /* أزرار الإجراءات */
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
        background-color: #1e3c72;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

excel_file = "Delta_Nile_HSE_Management_System_V2 (Recovered).xlsm"

# دالة تنظيف البيانات والهيدر معالجة ذكياً
def clean_dataframe(df):
    df = df.dropna(how='all').dropna(axis=1, how='all')
    header_idx = 0
    for idx, row in df.iterrows():
        if row.dropna().count() > 2:
            header_idx = idx
            break
            
    if header_idx > 0:
        new_header = df.iloc[header_idx].astype(str).str.strip()
        df = df.iloc[header_idx + 1:].copy()
        df.columns = new_header

    clean_cols = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if "Unnamed" in col_str or col_str in ["None", "nan", ""]:
            clean_cols.append(f"عمود_{i+1}")
        else:
            clean_cols.append(col_str)
    df.columns = clean_cols
    return df.replace({np.nan: "", "None": "", "nan": ""})

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
    
    # القائمة الجانبية للتنقل
    st.sidebar.markdown("<h2 style='text-align: center; color: #1e3c72;'>🛡️ لوحة التحكّم</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    selected_sheet = st.sidebar.selectbox("📂 اختر السجل / الشيت المرد عرضه:", list(data_sheets.keys()))
    
    df = data_sheets[selected_sheet]
    
    # الهيدر الرئيسي
    st.markdown(f"<h1 style='color: #1e3c72;'>📊 نظام السلامة والصحة المهنية - {selected_sheet}</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    # كروت مؤشرات الأداء المتطورة (KPIs)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"<div class='metric-card'><h3>إجمالي السجلات</h3><h2>{len(df)}</h2></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='metric-card'><h3>عدد أعمدة البيانات</h3><h2>{len(df.columns)}</h2></div>", unsafe_allow_html=True)
    with col3:
        # حساب نسبة الخلايا المكتملة
        total_cells = df.size
        filled_cells = df.astype(str).replace("", np.nan).count().sum()
        compliance_rate = round((filled_cells / total_cells) * 100) if total_cells > 0 else 0
        st.markdown(f"<div class='metric-card'><h3>نسبة اكتمال البيانات</h3><h2>{compliance_rate}%</h2></div>", unsafe_allow_html=True)
    with col4:
        st.markdown(f"<div class='metric-card'><h3>إجمالي السجلات بالنظام</h3><h2>{len(data_sheets)} شيت</h2></div>", unsafe_allow_html=True)

    st.write("")
    
    # تبويبات العمليات (عرض البيانات - رسوم بيانية - تعديل البيانات - تصدير)
    tab1, tab2, tab3, tab4 = st.tabs(["📋 عرض البيانات والتصفية", "📈 الرسوم البيانية والتحليلات", "✏️ تعديل وإضافة بيانات", "📥 تصدير التقارير"])
    
    with tab1:
        col_search, col_filter = st.columns([2, 1])
        with col_search:
            search_query = st.text_input("🔍 بحث شامل داخل السجل الحالي:")
        with col_filter:
            selected_col = st.selectbox("🎯 تصفية حسب العمود:", ["الكل"] + list(df.columns))
            
        filtered_df = df.copy()
        if selected_col != "الكل":
            filter_val = st.text_input(f"قيمة الفلتر لـ ({selected_col}):")
            if filter_val:
                filtered_df = filtered_df[filtered_df[selected_col].astype(str).str.contains(filter_val, case=False)]
                
        if search_query:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
            filtered_df = filtered_df[mask]
            
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        
    with tab2:
        st.subheader("📊 تحليلات ورسوم بيانية أوتوماتيكية")
        col_chart1, col_chart2 = st.columns(2)
        
        # محاولة رسم مخطط تناسبي لأول عمود نصي يحتوي على تكرارات
        categorical_cols = [col for col in df.columns if df[col].nunique() < 20 and df[col].nunique() > 1]
        if categorical_cols:
            target_col = categorical_cols[0]
            val_counts = df[target_col].value_counts().reset_index()
            val_counts.columns = [target_col, 'العدد']
            
            with col_chart1:
                fig_bar = px.bar(val_counts, x=target_col, y='العدد', title=f"توزيع البيانات حسب: {target_col}", color='العدد', color_continuous_scale='Blues')
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with col_chart2:
                fig_pie = px.pie(val_counts, names=target_col, values='العدد', title=f"نسب توزيع: {target_col}", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("لا توجد أعمدة تكرارية كافية وإنشاء رسوم بيانية تلقائية لهذا الشيت.")
            
    with tab3:
        st.subheader("✏️ محرر البيانات التفاعلي (يمكنك التعديل مباشرة في الجدول)")
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="data_editor")
        if st.button("💾 حفظ التغييرات مؤقتاً"):
            st.success("تم تحديث البيانات في الجلسة الحالية بنجاح!")
            
    with tab4:
        st.subheader("📥 تحميل البيانات والتقارير")
        col_down1, col_down2 = st.columns(2)
        
        # تصدير CSV
        csv_data = filtered_df.to_csv(index=False).encode('utf-8-sig')
        col_down1.download_button(
            label="📄 تحميل التقرير الحالي بصيغة (CSV)",
            data=csv_data,
            file_name=f"{selected_sheet}.csv",
            mime="text/csv"
        )

except Exception as e:
    st.error(f"حدث خطأ أثناء تحميل لوحة التحكم: {e}")
