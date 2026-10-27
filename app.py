import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="منظومة التقارير الهندسية والتقييم العقاري", layout="wide")

st.title("🏗️ نظام توليد التقارير الهندسية وتقييم العقارات الآلي")
st.markdown("---")

# إدخال البيانات الأساسية للتقرير
st.sidebar.header("📋 بيانات العقار والتقرير")
report_id = st.sidebar.text_input("رقم القضية / التقرير", "Case 137, 1447H")
engineer_name = st.sidebar.text_input("اسم الخبير / المهندس", "المهندس المقيم")
location = st.sidebar.text_input("موقع العقار", "المملكة العربية السعودية")

st.sidebar.markdown("---")
st.sidebar.header("📐 المدخلات الهندسية والمساحية")
land_area = st.sidebar.number_input("مساحة الأرض (متر مربع)", value=500.0, step=10.0)
land_unit_price = st.sidebar.number_input("سعر المتر المسطح للأرض", value=1500.0, step=50.0)

building_area = st.sidebar.number_input("مسطح البناء الإجمالي (متر مربع)", value=800.0, step=20.0)
building_unit_cost = st.sidebar.number_input("تكلفة بناء المتر المسطح (الجديدة)", value=2200.0, step=50.0)
building_age = st.sidebar.number_input("عمر المبنى (سنوات)", value=5, step=1)
depreciation_rate = st.sidebar.slider("نسبة الإهلاك السنوي (%)", 0.0, 5.0, 1.0)

# الحسابات الهندسية والمالية
total_land_value = land_area * land_unit_price
replacement_cost = building_area * building_unit_cost
total_depreciation = replacement_cost * (depreciation_rate / 100.0) * building_age
net_building_value = max(0.0, replacement_cost - total_depreciation)
total_property_value = total_land_value + net_building_value

# عرض النتائج في الواجهة
st.subheader(f"📊 نتائج التقييم المالي والهندسي - التقرير: {report_id}")

col1, col2, col3 = st.columns(3)
col1.metric("إجمالي قيمة الأرض", f"{total_land_value:,.2f} ر.س")
col2.metric("القيمة الحالية للمباني", f"{net_building_value:,.2f} ر.س")
col3.metric("إجمالي قيمة العقار التقديرية", f"{total_property_value:,.2f} ر.س")

st.markdown("---")

# جدول تفصيلي للكميات والتقييم
summary_data = {
    "البيان الهندسري / العنصر": [
        "قيمة الأرض (مساحة × سعر المتر)",
        "تكلفة الإنشاء الجديدة للمباني",
        "خصم الإهلاك المتراكم",
        "القيمة الدفترية الحالية للمباني",
        "القيمة الإجمالية النهائية للعقار"
    ],
    "القيمة / الكمية": [
        f"{total_land_value:,.2f} ر.س",
        f"{replacement_cost:,.2f} ر.س",
        f"-{total_depreciation:,.2f} ر.س",
        f"{net_building_value:,.2f} ر.س",
        f"{total_property_value:,.2f} ر.س"
    ]
}

df_summary = pd.DataFrame(summary_data)
st.table(df_summary)

# إنشاء ملف إكسل احترافي متعدد الجداول والبيانات
def create_excel_report():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_summary.to_excel(writer, sheet_name='ملخص التقييم', index=False)
        
        # جدول تفصيلي إضافي للمساحات
        details_data = {
            "المتغير الهندسي": ["مساحة الأرض", "مسطح البناء", "عمر المبنى", "معدل الإهلاك"],
            "القيمة": [land_area, building_area, building_age, f"{depreciation_rate}%"]
        }
        df_details = pd.DataFrame(details_data)
        df_details.to_excel(writer, sheet_name='البيانات والمساحات', index=False)
        
    output.seek(0)
    return output

excel_data = create_excel_report()

st.download_button(
    label="📥 تحميل التقرير بصيغة Excel معتمد",
    data=excel_data,
    file_name=f"Engineering_Report_{report_id.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
