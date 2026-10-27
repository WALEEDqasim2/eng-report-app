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
location = st.sidebar.text_input("موقع العقار", "الجمهورية اليمنية")

st.sidebar.markdown("---")
st.sidebar.header("📐 المدخلات الهندسية والمساحية")
land_area = st.sidebar.number_input("مساحة الأرض (متر مربع)", value=500.0, step=10.0)
land_unit_price = st.sidebar.number_input("سعر المتر المسطح للأرض (بالريال اليمني)", value=45000.0, step=1000.0)

building_area = st.sidebar.number_input("مسطح البناء الإجمالي (متر مربع)", value=800.0, step=20.0)
building_unit_cost = st.sidebar.number_input("تكلفة بناء المتر المسطح الجديدة (بالريال اليمني)", value=65000.0, step=1000.0)
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
col1.metric("إجمالي قيمة الأرض", f"{total_land_value:,.2f} ر.ي")
col2.metric("القيمة الحالية للمباني", f"{net_building_value:,.2f} ر.ي")
col3.metric("إجمالي قيمة العقار التقديرية", f"{total_property_value:,.2f} ر.ي")

st.markdown("---")

# جدول الخلاصة للعرض على الشاشة
summary_data = {
    "البيان الهندسي / العنصر": [
        "قيمة الأرض (مساحة × سعر المتر)",
        "تكلفة الإنشاء الجديدة للمباني",
        "خصم الإهلاك المتراكم",
        "القيمة الدفترية الحالية للمباني",
        "القيمة الإجمالية النهائية للعقار"
    ],
    "القيمة / الكمية (ر.ي)": [
        total_land_value,
        replacement_cost,
        -total_depreciation,
        net_building_value,
        total_property_value
    ]
}
df_summary = pd.DataFrame(summary_data)
st.table(df_summary)

# دالة توليد ملف Excel احترافي متعدد الشيتات والجداول المنسقة والألوان باستخدام xlsxwriter
def create_professional_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # تنسيقات الألوان الاحترافية
        header_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#1F4E78',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_color': '#1F4E78',
            'align': 'center'
        })
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })
        text_cell_format = workbook.add_format({
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        total_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9E1F2',
            'border': 1,
            'align': 'right',
            'num_format': '#,##0.00'
        })

        # --- الشيت الأول: ملخص التقييم المالي ---
        ws1 = workbook.add_worksheet('ملخص التقييم المالي')
        ws1.right_to_left()
        
        ws1.write(0, 1, f"تقرير تقييم عقاري - رقم القضية: {report_id}", title_format)
        ws1.write(1, 1, f"المهندس الخبير: {engineer_name} | الموقع: {location}", workbook.add_format({'align': 'center', 'italic': True}))
        
        headers = ["البيان الهندسي / العنصر", "القيمة / الكمية (ر.ي)"]
        for col_num, header in enumerate(headers):
            ws1.write(3, col_num, header, header_format)
            
        row_num = 4
        for i, row in df_summary.iterrows():
            ws1.write(row_num, 0, row["البيان الهندسي / العنصر"], text_cell_format)
            if i == 4:
                ws1.write(row_num, 1, row["القيمة / الكمية (ر.ي)"], total_format)
            else:
                ws1.write(row_num, 1, row["القيمة / الكمية (ر.ي)"], cell_format)
            row_num += 1

        # --- الشيت الثاني: التفاصيل الهندسية الكاملة ---
        ws2 = workbook.add_worksheet('البيانات والمساحات الهندسية')
        ws2.right_to_left()
        
        ws2.write(0, 1, "جدول المدخلات والمواصفات الهندسية التفصيلية", title_format)
        
        details_headers = ["المتغير الهندسي / المواصفات", "القيمة المسجلة", "الوحدة"]
        for col_num, header in enumerate(details_headers):
            ws2.write(2, col_num, header, header_format)
            
        details_rows = [
            ["مساحة الأرض", land_area, "متر مربع"],
            ["سعر المتر المسطح للأرض", land_unit_price, "ر.ي / م²"],
            ["إجمالي قيمة الأرض", total_land_value, "ر.ي"],
            ["مسطح البناء الإجمالي", building_area, "متر مربع"],
            ["تكلفة بناء المتر الجديدة", building_unit_cost, "ر.ي / م²"],
            ["تكلفة الإنشاء الكلية للمباني", replacement_cost, "ر.ي"],
            ["عمر المبنى", building_age, "سنوات"],
            ["نسبة الإهلاك السنوي", depreciation_rate, "%"],
            ["إجمالي الإهلاك المتراكم", total_depreciation, "ر.ي"],
            ["القيمة الدفترية الحالية للمباني", net_building_value, "ر.ي"],
            ["القيمة الإجمالية النهائية للعقار", total_property_value, "ر.ي"]
        ]
        
        for r_idx, r_data in enumerate(details_rows):
            ws2.write(3 + r_idx, 0, r_data[0], text_cell_format)
            ws2.write(3 + r_idx, 1, r_data[1], cell_format if isinstance(r_data[1], (int, float)) else text_cell_format)
            ws2.write(3 + r_idx, 2, r_data[2], text_cell_format)

        # ضبط مقاسات الأعمدة لتكون منسقة تلقائياً
        ws1.set_column('A', 45)
        ws1.set_column('B', 30)
        ws2.set_column('A', 40)
        ws2.set_column('B', 25)
        ws2.set_column('C', 18)

    output.seek(0)
    return output

excel_data = create_professional_excel()

st.download_button(
    label="📥 تحميل التقرير الاحترافي بصيغة Excel (شيتات متعددة + ألوان وتنسيقات هندسية)",
    data=excel_data,
    file_name=f"Engineering_Report_{report_id.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
