"""
نماذج البيانات - Data Models
هياكل البيانات المستخدمة في النظام
"""

from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime

@dataclass
class CompanySettings:
    """إعدادات الشركة"""
    company_name: str = ""
    company_address: str = ""
    company_phone: str = ""
    company_email: str = ""
    tax_number: str = ""
    currency: str = "ر.س"
    tax_rate: float = 15.0  # نسبة الضريبة الافتراضية
    date_format: str = "dd/mm/yyyy"
    tax_inclusive: bool = False
    financial_year_start: str = "يناير"
    country: str = "السعودية"

@dataclass
class Customer:
    """نموذج العميل"""
    customer_id: str
    name: str
    phone: str
    email: str
    address: str
    balance: float = 0.0
    credit_limit: float = 0.0
    tax_number: str = ""
    status: str = "نشط"

@dataclass
class Supplier:
    """نموذج المورد"""
    supplier_id: str
    name: str
    phone: str
    email: str
    address: str
    balance: float = 0.0
    tax_number: str = ""
    status: str = "نشط"

@dataclass
class Product:
    """نموذج المنتج"""
    product_id: str
    name: str
    category: str
    unit: str
    purchase_price: float
    sale_price: float
    opening_stock: float = 0.0
    current_stock: float = 0.0
    min_stock: float = 0.0
    status: str = "نشط"

@dataclass
class SalesInvoice:
    """نموذج فاتورة المبيعات"""
    invoice_number: str
    date: datetime
    customer_id: str
    items: List[dict]  # قائمة بالمنتجات والكميات والأسعار
    discount: float = 0.0
    tax: float = 0.0
    total: float = 0.0
    payment_method: str = "نقدي"
    status: str = "مكتمل"
    notes: str = ""

@dataclass
class PurchaseInvoice:
    """نموذج فاتورة المشتريات"""
    invoice_number: str
    date: datetime
    supplier_id: str
    items: List[dict]
    discount: float = 0.0
    tax: float = 0.0
    total: float = 0.0
    payment_method: str = "نقدي"
    status: str = "مكتمل"
    notes: str = ""

@dataclass
class Payment:
    """نموذج المدفوعات"""
    payment_id: str
    date: datetime
    type: str  # "دفعة عميل" أو "دفعة لمورد"
    party_id: str  # رقم العميل أو المورد
    amount: float
    method: str
    reference: str
    notes: str = ""

@dataclass
class Expense:
    """نموذج المصروفات"""
    expense_id: str
    date: datetime
    category: str
    description: str
    amount: float
    payment_method: str
    notes: str = ""