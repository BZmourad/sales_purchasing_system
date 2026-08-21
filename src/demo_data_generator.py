"""
مولد البيانات التجريبية - Demo Data Generator
إنشاء بيانات واقعية للعرض التجريبي
"""

import random
from datetime import datetime, timedelta
from typing import List, Dict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from constants import COLORS, FONT_STYLES, FILL_PATTERNS, ALIGNMENTS, BORDERS

class DemoDataGenerator:
    """منشئ البيانات التجريبية"""
    
    def __init__(self):
        self.customers = []
        self.suppliers = []
        self.products = []
        self.sales_invoices = []
        self.purchase_invoices = []
        self.payments = []
        self.expenses = []
        
        # توليد البيانات
        self.generate_customers()
        self.generate_suppliers()
        self.generate_products()
        self.generate_sales_invoices()
        self.generate_purchase_invoices()
        self.generate_payments()
        self.generate_expenses()
    
    def generate_customers(self):
        """توليد بيانات العملاء"""
        arabic_names = [
            'شركة الأمل للتجارة', 'مؤسسة النور', 'شركة المستقبل', 'مؤسسة البركة',
            'شركة السلام', 'مؤسسة الاتحاد', 'شركة الرواد', 'مؤسسة الإخاء',
            'شركة التميز', 'مؤسسة الفجر', 'شركة الواحة', 'مؤسسة الخليج',
            'شركة العربية', 'مؤسسة النخبة', 'شركة الإبداع', 'مؤسسة الصفا',
            'شركة الريادة', 'مؤسسة الدانة', 'شركة الأفق', 'مؤسسة الوئام'
        ]
        
        cities = ['الرياض', 'جدة', 'الدمام', 'الدوحة', 'الكويت', 'المنامة', 'القاهرة', 'الإسكندرية']
        
        for i in range(1, 21):
            customer = {
                'id': f'C{str(i).zfill(3)}',
                'name': arabic_names[i-1],
                'phone': f'05{random.randint(10000000, 99999999)}',
                'email': f'customer{i}@company.com',
                'address': f'{random.choice(cities)} - شارع الملك فهد - عمارة {random.randint(1, 50)}',
                'balance': round(random.uniform(10000, 150000), 2),
                'credit_limit': round(random.uniform(50000, 300000), 2),
                'tax_number': f'30{random.randint(100000000, 999999999)}',
                'status': random.choice(['نشط', 'نشط', 'نشط', 'معلق'])
            }
            self.customers.append(customer)
    
    def generate_suppliers(self):
        """توليد بيانات الموردين"""
        supplier_names = [
            'مصنع الإنتاج الأول', 'شركة التوريدات المتحدة', 'مؤسسة المواد الخام',
            'شركة الاستيراد والتصدير', 'مصنع البلاستيك الحديث', 'شركة الورق والكرتون',
            'مؤسسة المعادن', 'شركة الكيماويات', 'مصنع الأغذية', 'شركة الإلكترونيات',
            'مؤسسة التعبئة والتغليف', 'شركة النقل والتوزيع'
        ]
        
        for i in range(1, 13):
            supplier = {
                'id': f'S{str(i).zfill(3)}',
                'name': supplier_names[i-1],
                'phone': f'05{random.randint(10000000, 99999999)}',
                'email': f'supplier{i}@factory.com',
                'address': f'المدينة الصناعية - قطعة {random.randint(1, 100)}',
                'balance': round(random.uniform(20000, 200000), 2),
                'tax_number': f'30{random.randint(100000000, 999999999)}',
                'status': 'نشط'
            }
            self.suppliers.append(supplier)
    
    def generate_products(self):
        """توليد بيانات المنتجات"""
        products_data = [
            ('P001', 'حاسوب محمول', 'إلكترونيات', 'قطعة', 2500, 3200),
            ('P002', 'طابعة ليزر', 'إلكترونيات', 'قطعة', 800, 1100),
            ('P003', 'شاشة عرض 24 بوصة', 'إلكترونيات', 'قطعة', 600, 850),
            ('P004', 'لوحة مفاتيح', 'إلكترونيات', 'قطعة', 50, 90),
            ('P005', 'فأرة لاسلكية', 'إلكترونيات', 'قطعة', 40, 75),
            ('P006', 'ورق طباعة A4', 'قرطاسية', 'علبة', 25, 40),
            ('P007', 'أقلام حبر', 'قرطاسية', 'علبة', 15, 25),
            ('P008', 'ملفات بلاستيكية', 'قرطاسية', 'قطعة', 5, 10),
            ('P009', 'كرسي مكتب', 'أثاث', 'قطعة', 350, 550),
            ('P010', 'طاولة اجتماعات', 'أثاث', 'قطعة', 1200, 1800),
            ('P011', 'خزانة ملفات', 'أثاث', 'قطعة', 450, 700),
            ('P012', 'مكيف هواء', 'أجهزة', 'قطعة', 1500, 2200),
            ('P013', 'ثلاجة عرض', 'أجهزة', 'قطعة', 2800, 3800),
            ('P014', 'غسالة ملابس', 'أجهزة', 'قطعة', 1800, 2500),
            ('P015', 'مكنسة كهربائية', 'أجهزة', 'قطعة', 400, 650),
        ]
        
        for prod_id, name, category, unit, purchase_price, sale_price in products_data:
            opening_stock = random.randint(20, 100)
            product = {
                'id': prod_id,
                'name': name,
                'category': category,
                'unit': unit,
                'purchase_price': purchase_price,
                'sale_price': sale_price,
                'opening_stock': opening_stock,
                'current_stock': opening_stock,
                'min_stock': random.randint(5, 20),
                'status': 'نشط'
            }
            self.products.append(product)
    
    def generate_sales_invoices(self):
        """توليد فواتير المبيعات"""
        start_date = datetime(2024, 1, 1)
        
        for i in range(1, 51):
            customer = random.choice(self.customers)
            num_items = random.randint(1, 5)
            items = []
            total = 0
            
            for _ in range(num_items):
                product = random.choice(self.products)
                quantity = random.randint(1, 10)
                price = product['sale_price']
                item_total = quantity * price
                total += item_total
                items.append({
                    'product': product['name'],
                    'quantity': quantity,
                    'price': price,
                    'total': item_total
                })
            
            discount = round(total * random.uniform(0, 0.1), 2)
            tax = round((total - discount) * 0.15, 2)
            final_total = total - discount + tax
            
            invoice = {
                'number': f'INV-{str(i).zfill(4)}',
                'date': start_date + timedelta(days=random.randint(0, 220)),
                'customer': customer['name'],
                'items': items,
                'total_before_discount': total,
                'discount': discount,
                'tax': tax,
                'final_total': final_total,
                'payment_method': random.choice(['نقدي', 'تحويل بنكي', 'آجل']),
                'status': 'مكتمل'
            }
            self.sales_invoices.append(invoice)
    
    def generate_purchase_invoices(self):
        """توليد فواتير المشتريات"""
        start_date = datetime(2024, 1, 1)
        
        for i in range(1, 31):
            supplier = random.choice(self.suppliers)
            num_items = random.randint(1, 4)
            items = []
            total = 0
            
            for _ in range(num_items):
                product = random.choice(self.products)
                quantity = random.randint(10, 50)
                price = product['purchase_price']
                item_total = quantity * price
                total += item_total
                items.append({
                    'product': product['name'],
                    'quantity': quantity,
                    'price': price,
                    'total': item_total
                })
            
            discount = round(total * random.uniform(0, 0.05), 2)
            tax = round((total - discount) * 0.15, 2)
            final_total = total - discount + tax
            
            invoice = {
                'number': f'PO-{str(i).zfill(4)}',
                'date': start_date + timedelta(days=random.randint(0, 220)),
                'supplier': supplier['name'],
                'items': items,
                'total_before_discount': total,
                'discount': discount,
                'tax': tax,
                'final_total': final_total,
                'payment_method': random.choice(['نقدي', 'تحويل بنكي', 'شيك']),
                'status': 'مكتمل'
            }
            self.purchase_invoices.append(invoice)
    
    def generate_payments(self):
        """توليد المدفوعات"""
        start_date = datetime(2024, 1, 1)
        
        for i in range(1, 41):
            payment_type = random.choice(['دفعة عميل', 'دفعة لمورد'])
            if payment_type == 'دفعة عميل':
                party = random.choice(self.customers)['name']
                amount = round(random.uniform(5000, 50000), 2)
            else:
                party = random.choice(self.suppliers)['name']
                amount = round(random.uniform(10000, 80000), 2)
            
            payment = {
                'id': f'PAY-{str(i).zfill(4)}',
                'date': start_date + timedelta(days=random.randint(0, 220)),
                'type': payment_type,
                'party': party,
                'amount': amount,
                'method': random.choice(['نقدي', 'تحويل بنكي', 'شيك']),
                'reference': f'REF-{random.randint(10000, 99999)}',
                'notes': 'دفعة مستحقة'
            }
            self.payments.append(payment)
    
    def generate_expenses(self):
        """توليد المصروفات"""
        start_date = datetime(2024, 1, 1)
        expense_categories = ['إيجار', 'رواتب', 'كهرباء', 'ماء', 'اتصالات', 'أخرى']
        
        for i in range(1, 31):
            expense = {
                'id': f'EXP-{str(i).zfill(4)}',
                'date': start_date + timedelta(days=random.randint(0, 220)),
                'category': random.choice(expense_categories),
                'description': f'مصروف {i}',
                'amount': round(random.uniform(1000, 15000), 2),
                'payment_method': random.choice(['نقدي', 'تحويل بنكي']),
                'notes': ''
            }
            self.expenses.append(expense)

class DemoDataPopulator:
    """تعبئة البيانات التجريبية في Excel"""
    
    def __init__(self, generator):
        self.generator = generator
        self.demo_data = DemoDataGenerator()
    
    def populate_all_sheets(self):
        """تعبئة جميع الأوراق بالبيانات"""
        self.populate_dashboard()
        self.populate_settings()
        self.populate_customers()
        self.populate_suppliers()
        self.populate_products()
        self.populate_inventory()
        self.populate_sales()
        self.populate_purchases()
        self.populate_payments()
        self.populate_expenses()
    
    def populate_dashboard(self):
        """تعبئة لوحة التحكم"""
        ws = self.generator.sheets.get('dashboard')
        if not ws:
            return
        
        # حساب الإحصائيات
        total_sales = sum(inv['final_total'] for inv in self.demo_data.sales_invoices)
        total_purchases = sum(inv['final_total'] for inv in self.demo_data.purchase_invoices)
        gross_profit = total_sales - total_purchases
        outstanding_receivables = sum(c['balance'] for c in self.demo_data.customers)
        
        # تحديث القيم في البطاقات
        values = {
            3: [total_sales, total_purchases, len(self.demo_data.customers)],
            5: [len(self.demo_data.products), outstanding_receivables, gross_profit]
        }
        
        for row, row_values in values.items():
            for col_index, value in enumerate(row_values):
                col = 2 + (col_index * 2)  # B, D, F
                cell = ws.cell(row=row+1, column=col)
                cell.value = f"{value:,.2f}" if isinstance(value, float) else value
                cell.font = Font(name='Arial', size=14, bold=True, color=COLORS['primary'])
    
    def populate_settings(self):
        """تعبئة الإعدادات"""
        ws = self.generator.sheets.get('settings')
        if not ws:
            return
        
        settings_values = [
            'شركة النجاح للتجارة والتوزيع',
            'الرياض - حي العليا - شارع التحلية',
            '0112345678',
            'info@alnajah.com',
            '310123456789',
            'ر.س',
            15,
            'السعودية'
        ]
        
        for i, value in enumerate(settings_values, start=3):
            ws.cell(row=i, column=2, value=value)
    
    def populate_customers(self):
        """تعبئة العملاء"""
        ws = self.generator.sheets.get('customers')
        if not ws:
            return
        
        for i, customer in enumerate(self.demo_data.customers, start=3):
            ws.cell(row=i, column=1, value=customer['id'])
            ws.cell(row=i, column=2, value=customer['name'])
            ws.cell(row=i, column=3, value=customer['phone'])
            ws.cell(row=i, column=4, value=customer['email'])
            ws.cell(row=i, column=5, value=customer['address'])
            ws.cell(row=i, column=6, value=customer['balance'])
            ws.cell(row=i, column=7, value=customer['credit_limit'])
            ws.cell(row=i, column=8, value=customer['tax_number'])
            ws.cell(row=i, column=9, value=customer['status'])
    
    def populate_suppliers(self):
        """تعبئة الموردين"""
        ws = self.generator.sheets.get('suppliers')
        if not ws:
            return
        
        for i, supplier in enumerate(self.demo_data.suppliers, start=3):
            ws.cell(row=i, column=1, value=supplier['id'])
            ws.cell(row=i, column=2, value=supplier['name'])
            ws.cell(row=i, column=3, value=supplier['phone'])
            ws.cell(row=i, column=4, value=supplier['email'])
            ws.cell(row=i, column=5, value=supplier['address'])
            ws.cell(row=i, column=6, value=supplier['balance'])
            ws.cell(row=i, column=7, value=supplier['tax_number'])
            ws.cell(row=i, column=8, value=supplier['status'])
    
    def populate_products(self):
        """تعبئة المنتجات"""
        ws = self.generator.sheets.get('products')
        if not ws:
            return
        
        for i, product in enumerate(self.demo_data.products, start=3):
            ws.cell(row=i, column=1, value=product['id'])
            ws.cell(row=i, column=2, value=product['name'])
            ws.cell(row=i, column=3, value=product['category'])
            ws.cell(row=i, column=4, value=product['unit'])
            ws.cell(row=i, column=5, value=product['purchase_price'])
            ws.cell(row=i, column=6, value=product['sale_price'])
            ws.cell(row=i, column=7, value=product['opening_stock'])
            ws.cell(row=i, column=8, value=product['current_stock'])
            ws.cell(row=i, column=9, value=product['min_stock'])
            ws.cell(row=i, column=10, value=product['status'])
    
    def populate_inventory(self):
        """تعبئة المخزون"""
        ws = self.generator.sheets.get('inventory')
        if not ws:
            return
        
        for i, product in enumerate(self.demo_data.products, start=3):
            ws.cell(row=i, column=1, value=product['id'])
            ws.cell(row=i, column=2, value=product['name'])
            ws.cell(row=i, column=3, value=product['category'])
            ws.cell(row=i, column=4, value=product['unit'])
            ws.cell(row=i, column=5, value=product['opening_stock'])
            ws.cell(row=i, column=6, value=random.randint(10, 50))  # وارد
            ws.cell(row=i, column=7, value=random.randint(5, 30))   # صادر
            ws.cell(row=i, column=9, value=product['min_stock'])
            ws.cell(row=i, column=10, value=product['status'])
    
    def populate_sales(self):
        """تعبئة فواتير المبيعات"""
        ws = self.generator.wb['فواتير المبيعات']
        
        for i, invoice in enumerate(self.demo_data.sales_invoices, start=3):
            ws.cell(row=i, column=1, value=invoice['number'])
            ws.cell(row=i, column=2, value=invoice['date'].strftime('%d/%m/%Y'))
            ws.cell(row=i, column=3, value=invoice['customer'])
            ws.cell(row=i, column=4, value=invoice['items'][0]['product'])
            ws.cell(row=i, column=5, value=invoice['items'][0]['quantity'])
            ws.cell(row=i, column=6, value=invoice['items'][0]['price'])
            ws.cell(row=i, column=7, value=invoice['discount'])
            ws.cell(row=i, column=8, value=invoice['tax'])
            ws.cell(row=i, column=9, value=invoice['final_total'])
            ws.cell(row=i, column=10, value=invoice['payment_method'])
            ws.cell(row=i, column=11, value=invoice['status'])
    
    def populate_purchases(self):
        """تعبئة فواتير المشتريات"""
        ws = self.generator.wb['فواتير المشتريات']
        
        for i, invoice in enumerate(self.demo_data.purchase_invoices, start=3):
            ws.cell(row=i, column=1, value=invoice['number'])
            ws.cell(row=i, column=2, value=invoice['date'].strftime('%d/%m/%Y'))
            ws.cell(row=i, column=3, value=invoice['supplier'])
            ws.cell(row=i, column=4, value=invoice['items'][0]['product'])
            ws.cell(row=i, column=5, value=invoice['items'][0]['quantity'])
            ws.cell(row=i, column=6, value=invoice['items'][0]['price'])
            ws.cell(row=i, column=7, value=invoice['discount'])
            ws.cell(row=i, column=8, value=invoice['tax'])
            ws.cell(row=i, column=9, value=invoice['final_total'])
            ws.cell(row=i, column=10, value=invoice['payment_method'])
            ws.cell(row=i, column=11, value=invoice['status'])
    
    def populate_payments(self):
        """تعبئة المدفوعات"""
        ws = self.generator.sheets.get('payments')
        if not ws:
            return
        
        for i, payment in enumerate(self.demo_data.payments, start=3):
            ws.cell(row=i, column=1, value=payment['id'])
            ws.cell(row=i, column=2, value=payment['date'].strftime('%d/%m/%Y'))
            ws.cell(row=i, column=3, value=payment['type'])
            ws.cell(row=i, column=4, value=payment['party'])
            ws.cell(row=i, column=5, value=payment['amount'])
            ws.cell(row=i, column=6, value=payment['method'])
            ws.cell(row=i, column=7, value=payment['reference'])
            ws.cell(row=i, column=8, value=payment['notes'])
    
    def populate_expenses(self):
        """تعبئة المصروفات"""
        ws = self.generator.sheets.get('expenses')
        if not ws:
            return
        
        for i, expense in enumerate(self.demo_data.expenses, start=3):
            ws.cell(row=i, column=1, value=expense['id'])
            ws.cell(row=i, column=2, value=expense['date'].strftime('%d/%m/%Y'))
            ws.cell(row=i, column=3, value=expense['category'])
            ws.cell(row=i, column=4, value=expense['description'])
            ws.cell(row=i, column=5, value=expense['amount'])
            ws.cell(row=i, column=6, value=expense['payment_method'])
            ws.cell(row=i, column=7, value=expense['notes'])