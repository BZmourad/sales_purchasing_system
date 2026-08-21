"""
مولد ملف Excel - Excel Generator
الفئة الرئيسية لإنشاء ملف Excel
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from typing import Dict, List, Tuple
import os
from datetime import datetime

# استيراد مباشر بدون نقطة
from constants import COLORS, COLUMN_WIDTHS, FONT_STYLES, FILL_PATTERNS, ALIGNMENTS, BORDERS
from text_content import ARABIC_TEXT, DROPDOWN_OPTIONS

class ExcelGenerator:
    """منشئ ملف Excel"""
    
    def __init__(self):
        self.wb = Workbook()
        self.text = ARABIC_TEXT
        self.dropdowns = DROPDOWN_OPTIONS
        self.sheets = {}
        self.setup_workbook()
    
    def setup_workbook(self):
        """إعداد المصنف"""
        # حذف الورقة الافتراضية
        default_sheet = self.wb.active
        self.wb.remove(default_sheet)
        
        # إنشاء الأوراق
        self.create_dashboard_sheet()
        self.create_settings_sheet()
        self.create_customers_sheet()
        self.create_suppliers_sheet()
        self.create_products_sheet()
        self.create_inventory_sheet()
        self.create_payments_sheet()
        self.create_expenses_sheet()
        self.create_database_sheet()
        
        # تعيين اتجاه الأوراق للغة العربية
        for sheet in self.wb.worksheets:
            sheet.sheet_view.rightToLeft = True
            sheet.sheet_view.showGridLines = False
            sheet.sheet_properties.tabColor = COLORS['primary']
    
    def create_dashboard_sheet(self):
        """إنشاء لوحة التحكم"""
        ws = self.wb.create_sheet('لوحة التحكم')
        self.sheets['dashboard'] = ws
        
        # عنوان لوحة التحكم
        ws.merge_cells('A1:H1')
        title_cell = ws['A1']
        title_cell.value = 'لوحة التحكم الرئيسية'
        title_cell.font = FONT_STYLES['title']
        title_cell.alignment = ALIGNMENTS['center']
        title_cell.fill = FILL_PATTERNS['header']
        ws.row_dimensions[1].height = 50
        
        # بطاقات المؤشرات - مؤشرات في الصف 3
        indicators_row3 = [
            ('إجمالي المبيعات', 2, 0),   # العمود B
            ('إجمالي المشتريات', 4, 0),  # العمود D
            ('عدد العملاء', 6, 0)         # العمود F
        ]
        
        # بطاقات المؤشرات - مؤشرات في الصف 5
        indicators_row5 = [
            ('عدد المنتجات', 2, 0),       # العمود B
            ('المستحقات', 4, 0),          # العمود D
            ('إجمالي الربح', 6, 0)        # العمود F
        ]
        
        # إنشاء المؤشرات في الصف 3
        for label, col, value in indicators_row3:
            self.create_indicator_card(ws, 3, col, label, value)
        
        # إنشاء المؤشرات في الصف 5
        for label, col, value in indicators_row5:
            self.create_indicator_card(ws, 5, col, label, value)
        
        # إضافة أزرار التنقل
        self.add_navigation_buttons(ws)
    
    def create_indicator_card(self, ws, row, col, label, value):
        """إنشاء بطاقة مؤشر"""
        # دمج خليتين للعنوان
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+1)
        
        # كتابة العنوان
        title_cell = ws.cell(row=row, column=col)
        title_cell.value = label
        title_cell.font = FONT_STYLES['bold']
        title_cell.fill = FILL_PATTERNS['input']
        title_cell.alignment = ALIGNMENTS['center']
        title_cell.border = BORDERS['thin']
        
        # دمج خليتين للقيمة
        ws.merge_cells(start_row=row+1, start_column=col, end_row=row+1, end_column=col+1)
        
        # كتابة القيمة
        value_cell = ws.cell(row=row+1, column=col)
        value_cell.value = value
        value_cell.font = Font(name='Arial', size=14, bold=True, color=COLORS['primary'])
        value_cell.fill = FILL_PATTERNS['white']
        value_cell.alignment = ALIGNMENTS['center']
        value_cell.border = BORDERS['thin']
        
        # تعيين ارتفاع الصفوف
        ws.row_dimensions[row].height = 25
        ws.row_dimensions[row+1].height = 35
    
    def add_navigation_buttons(self, ws):
        """إضافة أزرار التنقل"""
        buttons = [
            ('العملاء', 8, 2),
            ('الموردين', 8, 4),
            ('المنتجات', 8, 6),
            ('المبيعات', 10, 2),
            ('المشتريات', 10, 4),
            ('التقارير', 10, 6)
        ]
        
        for text, row, col in buttons:
            self.add_button(ws, row, col, text)
    
    def get_cell_coordinates(self, cell):
        """الحصول على إحداثيات الخلية"""
        from openpyxl.utils import column_index_from_string
        col_letter = ''.join([c for c in cell if c.isalpha()])
        row_num = int(''.join([c for c in cell if c.isdigit()]))
        col_num = column_index_from_string(col_letter)
        return row_num, col_num
    
    def create_settings_sheet(self):
        """إنشاء ورقة الإعدادات"""
        ws = self.wb.create_sheet('الإعدادات')
        self.sheets['settings'] = ws
        
        # عنوان الورقة
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        title_cell.value = 'إعدادات النظام'
        title_cell.font = FONT_STYLES['title']
        title_cell.alignment = ALIGNMENTS['center']
        title_cell.fill = FILL_PATTERNS['header']
        ws.row_dimensions[1].height = 40
        
        # حقول الإعدادات
        settings_fields = [
            ('company_name', 'اسم الشركة', ''),
            ('company_address', 'عنوان الشركة', ''),
            ('company_phone', 'هاتف الشركة', ''),
            ('company_email', 'البريد الإلكتروني', ''),
            ('tax_number', 'الرقم الضريبي', ''),
            ('currency', 'العملة', 'ر.س'),
            ('tax_rate', 'نسبة الضريبة %', 15),
            ('country', 'الدولة', 'السعودية')
        ]
        
        for i, (key, label, default) in enumerate(settings_fields, start=3):
            # كتابة العنوان
            label_cell = ws.cell(row=i, column=1, value=label)
            label_cell.font = FONT_STYLES['bold']
            label_cell.alignment = ALIGNMENTS['right']
            label_cell.fill = FILL_PATTERNS['input']
            label_cell.border = BORDERS['thin']
            
            # كتابة القيمة
            value_cell = ws.cell(row=i, column=2, value=default)
            value_cell.font = FONT_STYLES['normal']
            value_cell.alignment = ALIGNMENTS['left']
            value_cell.fill = FILL_PATTERNS['white']
            value_cell.border = BORDERS['thin']
        
        # تعيين عرض الأعمدة
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
    
    def create_customers_sheet(self):
        """إنشاء ورقة العملاء"""
        ws = self.wb.create_sheet('العملاء')
        self.sheets['customers'] = ws
        
        headers = [
            'رقم العميل',
            'اسم العميل',
            'رقم الهاتف',
            'البريد الإلكتروني',
            'العنوان',
            'الرصيد',
            'حد الائتمان',
            'الرقم الضريبي',
            'الحالة'
        ]
        
        self.create_table_sheet(ws, headers, 'customers_table', 500)
        
        # إضافة التحقق من البيانات
        self.add_data_validation(ws, 'I3:I502', self.dropdowns['status'])
        
        # إضافة أزرار
        self.add_button(ws, 1, 1, 'إضافة جديد')
        self.add_button(ws, 1, 3, 'بحث')
        self.add_button(ws, 1, 5, 'طباعة')
    
    def create_suppliers_sheet(self):
        """إنشاء ورقة الموردين"""
        ws = self.wb.create_sheet('الموردين')
        self.sheets['suppliers'] = ws
        
        headers = [
            'رقم المورد',
            'اسم المورد',
            'رقم الهاتف',
            'البريد الإلكتروني',
            'العنوان',
            'الرصيد',
            'الرقم الضريبي',
            'الحالة'
        ]
        
        self.create_table_sheet(ws, headers, 'suppliers_table', 500)
        
        # إضافة التحقق من البيانات
        self.add_data_validation(ws, 'H3:H502', self.dropdowns['status'])
        
        # إضافة أزرار
        self.add_button(ws, 1, 1, 'إضافة جديد')
        self.add_button(ws, 1, 3, 'بحث')
        self.add_button(ws, 1, 5, 'طباعة')
    
    def create_products_sheet(self):
        """إنشاء ورقة المنتجات"""
        ws = self.wb.create_sheet('المنتجات')
        self.sheets['products'] = ws
        
        headers = [
            'كود المنتج',
            'اسم المنتج',
            'الفئة',
            'الوحدة',
            'سعر الشراء',
            'سعر البيع',
            'رصيد افتتاحي',
            'الرصيد الحالي',
            'الحد الأدنى',
            'الحالة'
        ]
        
        self.create_table_sheet(ws, headers, 'products_table', 1000)
        
        # إضافة التحقق من البيانات
        self.add_data_validation(ws, 'D3:D1002', self.dropdowns['units'])
        self.add_data_validation(ws, 'J3:J1002', self.dropdowns['status'])
        
        # إضافة أزرار
        self.add_button(ws, 1, 1, 'إضافة منتج')
        self.add_button(ws, 1, 3, 'بحث')
        self.add_button(ws, 1, 5, 'طباعة')
    
    def create_inventory_sheet(self):
        """إنشاء ورقة المخزون"""
        ws = self.wb.create_sheet('المخزون')
        self.sheets['inventory'] = ws
        
        headers = [
            'كود المنتج',
            'اسم المنتج',
            'الفئة',
            'الوحدة',
            'الرصيد الافتتاحي',
            'الوارد',
            'الصادر',
            'الرصيد الحالي',
            'الحد الأدنى',
            'الحالة'
        ]
        
        self.create_table_sheet(ws, headers, 'inventory_table', 500)
        
        # إضافة صيغة الرصيد الحالي
        for row in range(3, 503):
            cell = ws.cell(row=row, column=8, value=f'=E{row}+F{row}-G{row}')
            cell.fill = FILL_PATTERNS['calc']
        
        # إضافة تنسيق شرطي للمخزون المنخفض
        low_stock_rule = FormulaRule(
            formula=['$H3<$I3'],
            fill=PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),
            font=Font(color='FFFFFF', bold=True)
        )
        ws.conditional_formatting.add('A3:J503', low_stock_rule)
    
    def create_payments_sheet(self):
        """إنشاء ورقة المدفوعات"""
        ws = self.wb.create_sheet('المدفوعات')
        self.sheets['payments'] = ws
        
        headers = [
            'رقم الدفعة',
            'التاريخ',
            'النوع',
            'الطرف',
            'المبلغ',
            'طريقة الدفع',
            'المرجع',
            'ملاحظات'
        ]
        
        self.create_table_sheet(ws, headers, 'payments_table', 1000)
        
        # إضافة التحقق من البيانات
        self.add_data_validation(ws, 'C3:C1002', ['دفعة عميل', 'دفعة لمورد'])
        self.add_data_validation(ws, 'F3:F1002', self.dropdowns['payment_methods'])
    
    def create_expenses_sheet(self):
        """إنشاء ورقة المصروفات"""
        ws = self.wb.create_sheet('المصروفات')
        self.sheets['expenses'] = ws
        
        headers = [
            'رقم المصروف',
            'التاريخ',
            'الفئة',
            'الوصف',
            'المبلغ',
            'طريقة الدفع',
            'ملاحظات'
        ]
        
        self.create_table_sheet(ws, headers, 'expenses_table', 1000)
        
        # إضافة التحقق من البيانات
        self.add_data_validation(ws, 'C3:C1002', ['إيجار', 'رواتب', 'كهرباء', 'ماء', 'اتصالات', 'أخرى'])
        self.add_data_validation(ws, 'F3:F1002', self.dropdowns['payment_methods'])
    
    def create_database_sheet(self):
        """إنشاء ورقة قاعدة البيانات"""
        ws = self.wb.create_sheet('قاعدة البيانات')
        self.sheets['database'] = ws
        ws.sheet_state = 'hidden'  # إخفاء الورقة
        
        # هذه الورقة ستستخدم لتخزين البيانات المساعدة
        ws['A1'] = 'قاعدة البيانات المساعدة'
        ws['A1'].font = FONT_STYLES['title']
    
    def create_table_sheet(self, ws, headers, table_name, rows):
        """إنشاء ورقة جدول"""
        # كتابة العناوين في الصف 2
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col, value=header)
            cell.font = FONT_STYLES['header']
            cell.fill = FILL_PATTERNS['header']
            cell.alignment = ALIGNMENTS['center']
            cell.border = BORDERS['thin']
        
        # تعيين عرض الأعمدة
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20
        
        # تعيين ارتفاع الصف
        ws.row_dimensions[2].height = 30
        
        # تنسيق الخلايا الفارغة من الصف 3
        for row in range(3, rows + 3):
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = FONT_STYLES['normal']
                cell.alignment = ALIGNMENTS['center']
                cell.border = BORDERS['thin']
                if row % 2 == 0:
                    cell.fill = FILL_PATTERNS['alt_row']
    
    def add_data_validation(self, ws, cell_range, options, allow_blank=True):
        """إضافة التحقق من البيانات"""
        dv = DataValidation(
            type="list",
            formula1='"' + ','.join(options) + '"',
            allow_blank=allow_blank
        )
        dv.error = "يرجى اختيار قيمة من القائمة"
        dv.errorTitle = "قيمة غير صالحة"
        dv.prompt = "اختر من القائمة"
        dv.promptTitle = "التحقق من البيانات"
        ws.add_data_validation(dv)
        dv.add(cell_range)
    
    def add_button(self, ws, row, col, text):
        """إضافة زر"""
        # دمج خليتين للزر
        ws.merge_cells(start_row=row, start_column=col, 
                      end_row=row, end_column=col+1)
        
        # كتابة النص
        cell_obj = ws.cell(row=row, column=col)
        cell_obj.value = text
        cell_obj.font = FONT_STYLES['bold']
        cell_obj.fill = PatternFill(start_color=COLORS['accent'], 
                                   end_color=COLORS['accent'], 
                                   fill_type='solid')
        cell_obj.alignment = ALIGNMENTS['center']
        cell_obj.border = BORDERS['medium']
        
        # حماية الزر من التعديل
        cell_obj.protection = Protection(locked=True)
    
    def protect_sheet(self, ws, password='123456', editable_ranges=None):
        """حماية الورقة"""
        ws.protection.sheet = True
        ws.protection.password = password
    
    def save(self, filename, password='123456'):
        """
        حفظ الملف مع حماية بكلمة مرور
        
        Args:
            filename: مسار الملف
            password: كلمة المرور
        """
        # تعيين اتجاه جميع الأوراق
        for sheet in self.wb.worksheets:
            sheet.sheet_view.rightToLeft = True
        
        # حماية الأوراق من التعديل
        for sheet in self.wb.worksheets:
            sheet.protection.sheet = True
            sheet.protection.password = password
            sheet.protection.enable()
        
        # حفظ الملف بدون تشفير أولاً
        temp_file = filename + '.temp'
        self.wb.save(temp_file)
        
        # تشفير الملف بكلمة مرور
        try:
            from password_protection import protect_with_password
            protect_with_password(temp_file, password)
            
            # استبدال الملف الأصلي
            import shutil
            shutil.move(temp_file, filename)
            
            print(f"✅ تم حفظ الملف وحمايته بكلمة مرور")
        
        except ImportError:
            # إذا لم تكن المكتبة متوفرة، حفظ الملف بدون تشفير
            self.wb.save(filename)
            print("⚠️ تنبيه: لم يتم تثبيت مكتبة التشفير. تم حفظ الملف بدون كلمة مرور.")
            print("   لتثبيت المكتبة: pip install msoffcrypto-tool")