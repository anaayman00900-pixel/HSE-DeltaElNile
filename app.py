import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

# 1. إعدادات الصفحة الأساسية
st.set_page_config(
    page_title="نظام إدارة السلامة والصحة المهنية المتكامل",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. تصميم CSS متطور يدعم اللغة العربية (RTL)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
        direction: rtl;
        text-align: right;
    }
    
    .stApp {
        background-color: #f4f6f9;
    }
    
    .company-header {
        background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
        color: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 10px 20px rgba(0,0,0,0.12);
        margin-bottom: 25px;
        border-right: 8px solid #ff9900;
    }
    .company-header h1 { margin: 0; font-size: 26px; font-weight: 800; color: #ffffff; }
    .company-header p { margin: 5px 0 0 0; font-size: 15px; color: #dcdde1; }
    
    .kpi-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-top: 4px solid #1e3c72;
        text-align: center;
    }
    .kpi-card .title { font-size: 14px; color: #7f8c8d; font-weight: 600; }
    .kpi-card .value { font-size: 26px; color: #2c3e50; font-weight: 800; margin-top: 8px; }
    
    .stDataFrame {
        background: #ffffff;
        padding: 12px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

excel_file = "Delta_Nile_HSE_Management_System_V2 (Recovered).xlsm"

def clean_dataframe(df):
    df = df.dropna(how='all').dropna(axis=1, how='all')
    if df.empty:
        return df
        
    header_idx = 0
    for idx, row in df.iterrows():
        if row.dropna().count() >= 2:
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
    df = df.reset_index(drop=True)
    return df.replace({np.nan: "", "None": "", "nan": ""})

@st.cache_data
def load_and_structure_data(file_path):
    xls = pd.ExcelFile(file_path)
    
    delta_nile_sheets = {}
    tag_caps_sheets = {}
    all_sheets = {}
    
    for sheet_name in xls.sheet_names:
        try:
            raw_df = xls.parse(sheet_name)
            cleaned_df = clean_dataframe(raw_df)
            if cleaned_df.empty:
                continue
                
            all_sheets[sheet_name] = cleaned_df
            
            # تصنيف بناءً على محتوى الاسم
            if "تاج" in sheet_name or "tag" in sheet_name.lower():
                tag_caps_sheets[sheet_name] = cleaned_df
            else:
                delta_nile_sheets[sheet_name] = cleaned_df
        except Exception:
            continue
            
    # ضمان عدم فراغ أي قسم
    if not delta_nile_sheets:
        delta_nile_sheets = all_sheets
    if not tag_caps_sheets:
        tag_caps_sheets = all_sheets

    return {
        "شركة دلتا النيل للصناعات الغذائية": delta_nile_sheets,
        "شركة تاج كابس (TAG Caps)": tag_caps_sheets,
        "كافة السجلات والنماذج العامة": all_sheets
    }

try:
    all_data = load_and_structure_data(excel_file)
    
    st.sidebar.markdown("<h2 style='text-align: center; color: #1e3c72;'>⚙️ لوحة التحكم الإدارية</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    selected_company = st.sidebar.selectbox(
        "🏢 اختر الشركة المطلوبة:",
        list(all_data.keys()),
        index=0
    )
    
    company_sheets = all_data[selected_company]
    
    selected_sheet = st.sidebar.selectbox(
        "📋 اختر النموذج / السجل:",
        list(company_sheets.keys())
    )
    
    if f"df_{selected_company}_{selected_sheet}" not in st.session_state:
        st.session_state[f"df_{selected_company}_{selected_sheet}"] = company_sheets[selected_sheet].copy()
        
    current_df = st.session_state[f"df_{selected_company}_{selected_sheet}"]
    
    st.markdown(f"""
        <div class='company-header'>
            <h1>🛡️ نظام السلامة والصحة المهنية - {selected_company}</h1>
            <p>السجل المفتوح حالياً: <b>{selected_sheet}</b></p>
        </div>
    """, unsafe_allow_html=True)
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(f"<div class='kpi-card'><div class='title'>إجمالي عدد السجلات</div><div class='value'>{len(current_df)}</div></div>", unsafe_allow_html=True)
    with kpi2:
        st.markdown(f"<div class='kpi-card'><div class='title'>عدد أعمدة البيانات</div><div class='value'>{len(current_df.columns)}</div></div>", unsafe_allow_html=True)
    with kpi3:
        total_cells = current_df.size
        filled = current_df.astype(str).replace("", np.nan).dropna().count().sum()
        completion = round((filled / total_cells) * 100) if total_cells > 0 else 0
        st.markdown(f"<div class='kpi-card'><div class='title'>نسبة اكتمال البيانات</div><div class='value'>{completion}%</div></div>", unsafe_allow_html=True)
    with kpi4:
        st.markdown(f"<div class='kpi-card'><div class='title'>إجمالي النماذج المتاحة</div><div class='value'>{len(company_sheets)}</div></div>", unsafe_allow_html=True)

    st.write("")

    tab_view, tab_edit, tab_analytics, tab_export = st.tabs([
        "🔍 عرض وتصفية البيانات", 
        "✏️ محرر البيانات التفاعلي (تعديل/إضافة)", 
        "📊 الرسوم البيانية والتحليلات", 
        "📥 تصدير التقارير"
    ])

    with tab_view:
        st.subheader(f"عرض سجلات: {selected_sheet}")
        c_search, c_col = st.columns([2, 1])
        with c_search:
            search_word = st.text_input("🔍 بحث عام في السجل:")
        with c_col:
            filter_column = st.selectbox("🎯 تصفية بحسب عمود محدد:", ["جميع الأعمدة"] + list(current_df.columns))

        view_df = current_df.copy()
        if filter_column != "جميع الأعمدة":
            val = st.text_input(f"ادخل القيمة للبحث داخل ({filter_column}):")
            if val:
                view_df = view_df[view_df[filter_column].astype(str).str.contains(val, case=False, na=False)]

        if search_word:
            mask = view_df.astype(str).apply(lambda row: row.str.contains(search_word, case=False, na=False)).any(axis=1)
            view_df = view_df[mask]

        st.dataframe(view_df, use_container_width=True, hide_index=True)

    with tab_edit:
        st.subheader("✏️ محرر البيانات المباشر")
        st.info("💡 يمكنك تعديل الخانات مباشرة أو إضافة صفوف جديدة ثم الضغط على زر الحفظ.")
        
        edited_df = st.data_editor(
            current_df,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{selected_company}_{selected_sheet}"
        )
        
        if st.button("💾 حفظ التعديلات", type="primary"):
            st.session_state[f"df_{selected_company}_{selected_sheet}"] = edited_df
            st.success("✅ تم حفظ التعديلات بنجاح!")
            st.rerun()

    with tab_analytics:
        st.subheader("📊 الرسوم البيانية")
        col_graph1, col_graph2 = st.columns(2)
        
        cat_cols = [c for c in current_df.columns if current_df[c].nunique() < 15 and current_df[c].nunique() > 1]
        
        if cat_cols:
            selected_cat = st.selectbox("اختر العمود لتحليله بيانيًا:", cat_cols)
            counts = current_df[selected_cat].value_counts().reset_index()
            counts.columns = [selected_cat, "العدد"]
            
            with col_graph1:
                fig_bar = px.bar(counts, x=selected_cat, y="العدد", title=f"توزيع: {selected_cat}", color="العدد")
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with col_graph2:
                fig_pie = px.pie(counts, names=selected_cat, values="العدد", title=f"نسبة: {selected_cat}", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("ℹ️ لا تتوفر أعمدة تكرارية مناسبة لإنشاء رسم بياني لهذا السجل.")

    with tab_export:
        st.subheader("📥 تصدير السجلات")
        exp_col1, exp_col2 = st.columns(2)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            current_df.to_excel(writer, index=False, sheet_name=selected_sheet[:31])
        
        exp_col1.download_button(
            label="📗 تحميل التقرير (Excel .xlsx)",
            data=buffer.getvalue(),
            file_name=f"{selected_company}_{selected_sheet}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        csv_bytes = current_df.to_csv(index=False).encode('utf-8-sig')
        exp_col2.download_button(
            label="📄 تحميل التقرير (CSV)",
            data=csv_bytes,
            file_name=f"{selected_company}_{selected_sheet}.csv",
            mime="text/csv"
        )

except Exception as e:
    st.error(f"حدث خطأ في تحميل النظام: {e}")
