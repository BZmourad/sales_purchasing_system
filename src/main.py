"""
الملف الرئيسي - Main File
نقطة البداية لإنشاء النظام
"""

import os
import sys
from pathlib import Path

# إضافة المسار إلى النظام
sys.path.append(str(Path(__file__).parent))

from excel_generator import ExcelGenerator
from demo_data_generator import DemoDataPopulator
from constants import COLORS, FONT_STYLES, FILL_PATTERNS, ALIGNMENTS, BORDERS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import BarChart, PieChart, LineChart, Reference

def main():
    """الوظيفة الرئيسية"""
    print("بدء إنشاء نظام المبيعات والمشتريات...")
    print("=" * 50)
    
    # إنشاء المولد
    print("1. إنشاء هيكل المصنف...")
    generator = ExcelGenerator()
    
    # إضافة المزيد من الأوراق والتفاصيل
    print("2. إضافة أوراق المبيعات والمشتريات...")
    add_sales_sheet(generator)
    add_purchases_sheet(generator)
    add_reports_sheet(generator)
    
    # إضافة البيانات التجريبية
    print("3. توليد البيانات التجريبية...")
    populator = DemoDataPopulator(generator)
    populator.populate_all_sheets()
    
    # إنشاء مجلد الإخراج
    output_dir = Path(__file__).parent.parent / 'output'
    output_dir.mkdir(exist_ok=True)
    
    # حفظ الملف مع كلمة مرور
    print("4. حفظ الملف مع الحماية...")
    filename = output_dir / 'نظام_المبيعات_والمشتريات.xlsx'
    
    # استخدام كلمة المرور
    password = '123456'
    generator.save(str(filename), password)
    
    print("=" * 50)
    print(f"✅ تم إنشاء النظام بنجاح!")
    print(f"📁 موقع الملف: {filename}")
    print(f"🔐 كلمة المرور: {password}")
    print("=" * 50)
    
    return str(filename)

def add_charts(generator):
    """إضافة الرسوم البيانية"""
    # إضافة رسم بياني للوحة التحكم
    dashboard_ws = generator.sheets.get('dashboard')
    if dashboard_ws:
        # رسم بياني للمبيعات الشهرية
        chart = BarChart()
        chart.type = "col"
        chart.title = "المبيعات الشهرية"
        chart.y_axis.title = "المبيعات"
        chart.x_axis.title = "الشهر"
        
        data = Reference(dashboard_ws, min_col=2, min_row=15, max_col=3, max_row=20)
        cats = Reference(dashboard_ws, min_col=1, min_row=16, max_row=20)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.shape = 4
        dashboard_ws.add_chart(chart, "A12")
    
    # إضافة رسوم بيانية لورقة التقارير
    reports_ws = generator.sheets.get('reports')
    if reports_ws:
        # رسم بياني للمبيعات
        sales_chart = PieChart()
        sales_chart.title = "توزيع المبيعات"
        data = Reference(reports_ws, min_col=2, min_row=10, max_col=3, max_row=15)
        reports_ws.add_chart(sales_chart, "A12")
        
        # رسم بياني للمخزون
        inventory_chart = BarChart()
        inventory_chart.type = "bar"
        inventory_chart.title = "المخزون حسب الفئة"
        data = Reference(reports_ws, min_col=5, min_row=10, max_col=6, max_row=15)
        reports_ws.add_chart(inventory_chart, "E12")

def add_sales_sheet(generator):
    """إضافة ورقة المبيعات"""
    ws = generator.wb.create_sheet('فواتير المبيعات')
    generator.sheets['sales'] = ws
    
    # إعداد الأعمدة
    headers = [
        'رقم الفاتورة', 'التاريخ', 'العميل', 'المنتج', 'الكمية',
        'سعر الوحدة', 'الخصم', 'الضريبة', 'الإجمالي', 'طريقة الدفع',
        'الحالة', 'ملاحظات'
    ]
    
    generator.create_table_sheet(ws, headers, 'sales_table', 1000)
    
    # إضافة التحقق من البيانات
    generator.add_data_validation(ws, 'J3:J1002', generator.dropdowns['payment_methods'])
    generator.add_data_validation(ws, 'K3:K1002', generator.dropdowns['status'])
    
    # إضافة صيغ حسابية
    for row in range(3, 1003):
        ws.cell(row=row, column=8, value=f'=IF(G{row}="",0,(E{row}*F{row}-G{row})*0.15)')
        ws.cell(row=row, column=9, value=f'=IF(E{row}="",0,E{row}*F{row}-G{row}+H{row})')
    
    # تنسيق الخلايا
    for col_idx in [5, 6, 7, 8, 9]:
        for row in range(3, 1003):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = '#,##0.00'
            if col_idx in [8, 9]:
                cell.fill = FILL_PATTERNS['calc']
    
    # إضافة أزرار (بدون معامل الماكرو)
    generator.add_button(ws, 1, 1, 'إضافة فاتورة')
    generator.add_button(ws, 1, 3, 'بحث')
    generator.add_button(ws, 1, 5, 'طباعة')

def add_purchases_sheet(generator):
    """إضافة ورقة المشتريات"""
    ws = generator.wb.create_sheet('فواتير المشتريات')
    generator.sheets['purchases'] = ws
    
    headers = [
        'رقم الفاتورة', 'التاريخ', 'المورد', 'المنتج', 'الكمية',
        'سعر الشراء', 'الخصم', 'الضريبة', 'الإجمالي', 'طريقة الدفع',
        'الحالة', 'ملاحظات'
    ]
    
    generator.create_table_sheet(ws, headers, 'purchases_table', 1000)
    
    generator.add_data_validation(ws, 'J3:J1002', generator.dropdowns['payment_methods'])
    generator.add_data_validation(ws, 'K3:K1002', generator.dropdowns['status'])
    
    # إضافة صيغ حسابية
    for row in range(3, 1003):
        ws.cell(row=row, column=8, value=f'=IF(G{row}="",0,(E{row}*F{row}-G{row})*0.15)')
        ws.cell(row=row, column=9, value=f'=IF(E{row}="",0,E{row}*F{row}-G{row}+H{row})')
    
    # تنسيق الخلايا
    for col_idx in [5, 6, 7, 8, 9]:
        for row in range(3, 1003):
            cell = ws.cell(row=row, column=col_idx)
            cell.number_format = '#,##0.00'
            if col_idx in [8, 9]:
                cell.fill = FILL_PATTERNS['calc']
    
    # إضافة أزرار (بدون معامل الماكرو)
    generator.add_button(ws, 1, 1, 'إضافة فاتورة')
    generator.add_button(ws, 1, 3, 'بحث')
    generator.add_button(ws, 1, 5, 'طباعة')

def add_reports_sheet(generator):
    """إضافة ورقة التقارير"""
    ws = generator.wb.create_sheet('التقارير')
    generator.sheets['reports'] = ws
    
    # عنوان
    ws.merge_cells('A1:H1')
    title_cell = ws['A1']
    title_cell.value = 'مركز التقارير'
    title_cell.font = Font(name='Arial', size=18, bold=True, color=COLORS['primary'])
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color=COLORS['accent'], end_color=COLORS['accent'], fill_type='solid')
    ws.row_dimensions[1].height = 50
    
    # إضافة أزرار التقارير (بدون معامل الماكرو)
    reports = [
        ('تقرير المبيعات', 3, 2),
        ('تقرير المشتريات', 3, 4),
        ('تقرير المخزون', 3, 6),
        ('تقرير العملاء', 5, 2),
        ('تقرير الموردين', 5, 4),
        ('تقرير الأرباح والخسائر', 5, 6),
        ('تقرير الضرائب', 7, 2),
        ('تقرير المدفوعات', 7, 4)
    ]
    
    for text, row, col in reports:
        generator.add_button(ws, row, col, text)
if __name__ == '__main__':
    main()