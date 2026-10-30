import streamlit as st
import pandas as pd
import os

# 1. إعداد واجهة التطبيق وتفعيل التصميم العسكري الرسمي
st.set_page_config(page_title="منظومة تقييم الأداء البدني", layout="centered")

st.markdown("""
    <style>
    .stApp {
        background-color: #F4F6F4;
        direction: RTL;
        text-align: right;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stNumberInput"] label {
        text-align: right !important;
        width: 100% !important;
        display: block !important;
    }
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input {
        text-align: right !important;
        direction: rtl !important;
    }
    div[data-testid="stCheckbox"] {
        direction: rtl !important;
        text-align: right !important;
    }
    .centered-main-title {
        text-align: center;
        color: #2D3A22;
        font-weight: 900;
        font-size: 30px;
        margin-bottom: 5px;
        border-bottom: 3px solid #4A5D3B;
        padding-bottom: 10px;
    }
    .centered-main-subtitle {
        text-align: center;
        color: #556B2F;
        font-size: 15px;
        margin-bottom: 35px;
        font-weight: 500;
    }
    .centered-section-title {
        text-align: center;
        color: #2D3A22;
        font-weight: bold;
        font-size: 22px;
        margin-top: 25px;
        margin-bottom: 15px;
        background-color: #E2E8E2;
        padding: 8px;
        border-right: 5px solid #4A5D3B;
        border-left: 5px solid #4A5D3B;
    }
    .status-box {
        background-color: #E2E8E2;
        border-right: 4px solid #4A5D3B;
        padding: 10px;
        font-weight: bold;
        color: #2D3A22;
        text-align: center;
        margin-top: 25px;
    }
    .stButton>button {
        background-color: #3B4A2E !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 0px !important;
        border: 1px solid #2D3A22 !important;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #4A5D3B !important;
    }
    table {
        direction: rtl !important;
        text-align: right !important;
        width: 100%;
    }
    th, td {
        text-align: right !important;
        padding: 10px !important;
    }
    .metric-container {
        background-color: #E2E8E2;
        padding: 15px;
        text-align: center;
        font-size: 24px;
        font-weight: bold;
        color: #2D3A22;
        border: 1px solid #4A5D3B;
        margin-top: 20px;
        direction: ltr !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="centered-main-title">منظومة معايير الأداء البدني</h1>', unsafe_allow_html=True)
st.markdown('<p class="centered-main-subtitle">إدارة وحساب علامات قياس الكفاءة البدنية للمتدربين بناءً على مراجع قاعدة البيانات</p>', unsafe_allow_html=True)

# 2. دالة قراءة قاعدة البيانات
def load_fitness_data():
    file_name = "data.xlsx"
    if not os.path.exists(file_name) and os.path.exists("data.xlsx.xlsx"):
        file_name = "data.xlsx.xlsx"
        
    if os.path.exists(file_name):
        try:
            df = pd.read_excel(file_name, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
            # توحيد أسماء الأعمدة الأساسية
            df.columns = ['العدد', 'العلامة', 'الفئة العمرية', 'التمرين']
            df['التمرين'] = df['التمرين'].astype(str).str.strip()
            df['الفئة العمرية'] = df['الفئة العمرية'].astype(str).str.strip()
            return df
        except Exception as e:
            st.error(f"خطأ أثناء قراءة ملف البيانات: {e}")
            return None
    else:
        st.error("لم يتم العثور على ملف قاعدة البيانات المعتمد data.xlsx")
        return None

df_standards = load_fitness_data()

# 3. دالة حساب الفئة العمرية تلقائياً لعام 2026
def calculate_age_group(age):
    if age <= 25: return 1
    elif 26 <= age <= 30: return 2
    elif 31 <= age <= 35: return 3
    elif 36 <= age <= 40: return 4
    elif 41 <= age <= 45: return 5
    elif 46 <= age <= 50: return 6
    elif 51 <= age <= 55: return 7
    elif 56 <= age <= 60: return 8
    else: return 1

# 4. دالة المطابقة الحرفية الشاملة والمضمونة لمنع اختفاء النتائج
def lookup_score(exercise_type, group, value):
    if df_standards is None:
        return 0
    try:
        # فلترة بناءً على الفئة العمرية أولاً كقيمة نصية نظيفة
        df_filtered = df_standards[df_standards['الفئة العمرية'] == str(group)].copy()
        
        # البحث عن الكلمة المفتاحية في عمود التمرين بمرونة كاملة
        if exercise_type == "ضغط":
            df_exercise = df_filtered[df_filtered['التمرين'].str.contains("ضغط", na=False)]
        elif exercise_type == "100":
            df_exercise = df_filtered[df_filtered['التمرين'].str.contains("100", na=False)]
        elif exercise_type == "معدة":
            df_exercise = df_filtered[df_filtered['التمرين'].str.contains("معدة", na=False)]
        else: # 1 كم
            df_exercise = df_filtered[df_filtered['التمرين'].str.contains("1", na=False)]
            
        if df_exercise.empty:
            return 0
            
        # تحويل الأرقام للمقارنة الحسابية
        df_exercise['العدد'] = pd.to_numeric(df_exercise['العدد'], errors='coerce')
        df_exercise['العلامة'] = pd.to_numeric(df_exercise['العلامة'], errors='coerce')
        df_exercise = df_exercise.dropna(subset=['العدد', 'العلامة'])
        
        # البحث عن المطابقة التامة
        exact_match = df_exercise[df_exercise['العدد'] == value]
        if not exact_match.empty:
            return int(exact_match.iloc[0]['العلامة'])
            
        # إذا لم يجد تطابقاً تاماً، يجلب أقرب قيمة متوفرة في النطاق البيني
        idx = (df_exercise['العدد'] - value).abs().idxmin()
        return int(df_exercise.loc[idx, 'العلامة'])
    except:
        return 0

# 5. واجهة إدخال البيانات الرسمية المنسقة
if df_standards is not None:
    st.markdown('<div class="centered-section-title">بيانات المتدرب والملف الشخصي</div>', unsafe_allow_html=True)
    
    col_name, col_age, col_group = st.columns(3)
    with col_name:
        athlete_name = st.text_input("اسم المتدرب:", value="أحمد")
    with col_age:
        athlete_age = st.number_input("العمر الحالي (سنة):", min_value=15, max_value=65, value=22)
    with col_group:
        age_group = calculate_age_group(athlete_age)
        st.markdown(f'<div class="status-box">الفئة المستهدفة: الفئة ({age_group})</div>', unsafe_allow_html=True)

    st.divider()

    st.markdown('<div class="centered-section-title">نتائج الاختبارات البدنية الميدانية</div>', unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        run_100m = st.number_input("زمن جري 100 متر (ثواني وأجزاء):", min_value=0.0, format="%.2f", value=15.60)
    with c2:
        run_1km = st.number_input("زمن جري 1 كم (الدقائق والثواني ككسر):", min_value=0.0, format="%.2f", value=3.50)

    # تقريب حقول الضغط والمعدة مباشرة للأسفل
    cb1, cb2 = st.columns(2)
    with cb1:
        do_pushups = st.checkbox("تفعيل تمرين الضغط للمتدرب", value=True)
        pushups = st.number_input("عدد تكرارات تمرين الضغط المحققة:", min_value=0, value=40, disabled=not do_pushups)
    with cb2:
        do_situps = st.checkbox("تفعيل تمرين المعدة للمتدرب", value=True)
        situps = st.number_input("عدد تكرارات تمرين المعدة المحققة:", min_value=0, value=35, disabled=not do_situps)

    st.write("") 

    if st.button("احتساب الدرجات وإصدار بيان التقرير النهائي", type="primary", use_container_width=True):
        exercises_names = []
        results_logged = []
        scores_earned = []
        
        # 1. حساب جري 100 متر
        score_run100m = lookup_score("100", age_group, run_100m)
        exercises_names.append("جري 100 متر")
        results_logged.append(f"{run_100m} ثانية")
        scores_earned.append(score_run100m)
        
        # 2. حساب جري 1 كم
        score_run1km = lookup_score("1كم", age_group, run_1km)
        exercises_names.append("جري 1 كم")
        results_logged.append(f"{run_1km} دقيقة")
        scores_earned.append(score_run1km)
        
        # 3. حساب تمرين الضغط (إذا تم تفعيله)
        if do_pushups:
            score_pushups = lookup_score("ضغط", age_group, pushups)
            exercises_names.append("تمرين الضغط")
            results_logged.append(f"{pushups} تكرار")
            scores_earned.append(score_pushups)
            
        # 4. حساب تمرين المعدة (إذا تم تفعيله)
        if do_situps:
            score_situps = lookup_score("معدة", age_group, situps)
            exercises_names.append("تمرين المعدة")
            results_logged.append(f"{situps} تكرار")
            scores_earned.append(score_situps)
            
        # حساب المعدل العام الإجمالي بدقة وثبات
        total_average = sum(scores_earned) / len(scores_earned)
        
        st.success(f"تم احتساب النتيجة الإجمالية بناءً على قياس ({len(scores_earned)}) تمارين بنجاح.")
        
        st.markdown('<h3 class="centered-main-title" style="font-size: 22px; margin-top:20px; border-bottom: 2px solid #4A5D3B;">بيان عرض النتائج وتقييم الكفاءة البدنية</h3>', unsafe_allow_html=True)
        
        # بناء وعرض جدول التقرير النهائي الصارم لليمين
        report_data = {
            "نوع الاختبار البدني": exercises_names,
            "النتيجة المسجلة": results_logged,
            "العلامة المستحقة (من 100)": scores_earned
        }
        
        st.table(pd.DataFrame(report_data))
        
        # عرض المعدل النهائي في صندوق مخصص محمي من الانقلاب التنسيقي
        st.markdown(f'<div class="metric-container">المعدل العام النهائي للمتدرب: {total_average:.2f} / 100</div>', unsafe_allow_html=True)
