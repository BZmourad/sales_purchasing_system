"""
ملف النصوص العربية - Arabic Text Content
جميع النصوص العربية المستخدمة في النظام
"""

ARABIC_TEXT = {
    # أسماء الأوراق - Sheet Names
    "sheets": {
        "dashboard": "لوحة التحكم",
        "settings": "الإعدادات",
        "customers": "العملاء",
        "suppliers": "الموردين",
        "products": "المنتجات",
        "categories": "الفئات",
        "sales_quotes": "عروض الأسعار",
        "sales_invoices": "فواتير المبيعات",
        "sales_returns": "مرتجعات المبيعات",
        "purchase_orders": "أوامر الشراء",
        "purchase_invoices": "فواتير المشتريات",
        "purchase_returns": "مرتجعات المشتريات",
        "inventory": "المخزون",
        "payments": "المدفوعات",
        "expenses": "المصروفات",
        "reports": "التقارير",
        "database": "قاعدة البيانات"
    },
    
    # عناوين الأعمدة - Column Headers
    "headers": {
        # العملاء
        "customer_id": "رقم العميل",
        "customer_name": "اسم العميل",
        "customer_phone": "رقم الهاتف",
        "customer_email": "البريد الإلكتروني",
        "customer_address": "العنوان",
        "customer_balance": "الرصيد",
        "customer_credit_limit": "حد الائتمان",
        "customer_tax_number": "الرقم الضريبي",
        
        # الموردين
        "supplier_id": "رقم المورد",
        "supplier_name": "اسم المورد",
        "supplier_phone": "رقم الهاتف",
        "supplier_email": "البريد الإلكتروني",
        "supplier_address": "العنوان",
        "supplier_balance": "الرصيد",
        
        # المنتجات
        "product_id": "كود المنتج",
        "product_name": "اسم المنتج",
        "category": "الفئة",
        "unit": "الوحدة",
        "purchase_price": "سعر الشراء",
        "sale_price": "سعر البيع",
        "opening_stock": "رصيد افتتاحي",
        "current_stock": "الرصيد الحالي",
        "min_stock": "الحد الأدنى",
        
        # الفواتير
        "invoice_number": "رقم الفاتورة",
        "invoice_date": "التاريخ",
        "customer": "العميل",
        "supplier": "المورد",
        "product": "المنتج",
        "quantity": "الكمية",
        "unit_price": "سعر الوحدة",
        "discount": "الخصم",
        "tax": "الضريبة",
        "total": "الإجمالي",
        "payment_method": "طريقة الدفع",
        "status": "الحالة",
        "notes": "ملاحظات"
    },
    
    # أزرار - Buttons
    "buttons": {
        "add": "إضافة جديد",
        "save": "حفظ",
        "edit": "تعديل",
        "delete": "حذف",
        "search": "بحث",
        "clear": "مسح",
        "print": "طباعة",
        "export": "تصدير",
        "dashboard": "لوحة التحكم",
        "reports": "التقارير",
        "back": "رجوع",
        "main_menu": "القائمة الرئيسية",
        "refresh": "تحديث"
    },
    
    # رسائل الخطأ - Error Messages
    "errors": {
        "required_field": "هذا الحقل مطلوب",
        "invalid_number": "يرجى إدخال رقم صحيح",
        "invalid_date": "يرجى إدخال تاريخ صحيح",
        "duplicate_id": "الرقم موجود مسبقاً",
        "negative_value": "لا يمكن إدخال قيمة سالبة",
        "product_not_found": "المنتج غير موجود",
        "customer_not_found": "العميل غير موجود",
        "supplier_not_found": "المورد غير موجود",
        "insufficient_stock": "الكمية المتاحة غير كافية",
        "invalid_quantity": "يرجى إدخال كمية صحيحة"
    },
    
    # لوحة التحكم - Dashboard
    "dashboard": {
        "total_sales": "إجمالي المبيعات",
        "total_purchases": "إجمالي المشتريات",
        "total_customers": "عدد العملاء",
        "total_suppliers": "عدد الموردين",
        "total_products": "عدد المنتجات",
        "low_stock_items": "منتجات منخفضة المخزون",
        "outstanding_receivables": "المستحقات غير المسددة",
        "outstanding_payables": "المطلوبات غير المسددة",
        "monthly_sales": "المبيعات الشهرية",
        "gross_profit": "إجمالي الربح",
        "net_profit": "صافي الربح",
        "top_products": "أفضل المنتجات",
        "top_customers": "أفضل العملاء"
    },
    
    # الإعدادات - Settings
    "settings": {
        "company_name": "اسم الشركة",
        "company_address": "عنوان الشركة",
        "company_phone": "هاتف الشركة",
        "company_email": "البريد الإلكتروني",
        "tax_number": "الرقم الضريبي",
        "currency": "العملة",
        "currency_symbol": "رمز العملة",
        "tax_rate": "نسبة الضريبة",
        "date_format": "صيغة التاريخ",
        "tax_inclusive": "الأسعار شاملة الضريبة",
        "financial_year_start": "بداية السنة المالية"
    },
    
    # عناوين التقارير - Report Titles
    "reports": {
        "sales_report": "تقرير المبيعات",
        "purchases_report": "تقرير المشتريات",
        "inventory_report": "تقرير المخزون",
        "customer_report": "تقرير العملاء",
        "supplier_report": "تقرير الموردين",
        "profit_loss": "تقرير الأرباح والخسائر",
        "tax_report": "تقرير الضرائب",
        "payments_report": "تقرير المدفوعات",
        "receivables_report": "تقرير المستحقات",
        "payables_report": "تقرير المطلوبات"
    }
}

# قوائم الاختيار - Dropdown Lists
DROPDOWN_OPTIONS = {
    "payment_methods": ["نقدي", "شيك", "تحويل بنكي", "بطاقة ائتمان", "آجل"],
    "units": ["قطعة", "كيلو", "متر", "لتر", "صندوق", "علبة", "خدمة"],
    "status": ["نشط", "غير نشط", "معلق", "مكتمل", "ملغي"],
    "countries": ["السعودية", "قطر", "الكويت", "البحرين", "مصر"],
    "currencies": ["ر.س", "ر.ق", "د.ك", "د.ب", "ج.م", "$", "€"],
    "months": ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", 
               "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"],
    "years": [str(year) for year in range(2020, 2031)]
}