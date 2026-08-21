"""
الملف الثابت - Constants
القيم الثابتة المستخدمة في النظام
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ألوان النظام - System Colors
COLORS = {
    'primary': '1F4E78',      # أزرق داكن
    'secondary': '2E75B6',    # أزرق متوسط
    'accent': 'FFC000',       # ذهبي
    'success': '00B050',      # أخضر
    'danger': 'FF0000',       # أحمر
    'warning': 'FFC000',      # أصفر
    'info': '00B0F0',         # أزرق فاتح
    'light_gray': 'F2F2F2',   # رمادي فاتح
    'white': 'FFFFFF',        # أبيض
    'black': '000000',        # أسود
    'header_bg': '1F4E78',    # خلفية العناوين
    'alt_row': 'D9E2F3',      # صف متناوب
    'input_bg': 'FFF2CC',     # خلفية المدخلات
    'calc_bg': 'E2EFDA'       # خلفية الحسابات
}

# أبعاد الأعمدة - Column Widths
COLUMN_WIDTHS = {
    'id': 12,
    'name': 30,
    'phone': 15,
    'email': 25,
    'address': 40,
    'product': 30,
    'quantity': 12,
    'price': 15,
    'discount': 12,
    'tax': 12,
    'total': 18,
    'date': 15,
    'notes': 35,
    'status': 12
}

# تنسيقات الخطوط - Font Styles
FONT_STYLES = {
    'header': Font(name='Arial', size=11, bold=True, color='FFFFFF'),
    'title': Font(name='Arial', size=16, bold=True, color='1F4E78'),
    'subtitle': Font(name='Arial', size=12, bold=True, color='2E75B6'),
    'normal': Font(name='Arial', size=10, color='000000'),
    'bold': Font(name='Arial', size=10, bold=True, color='000000'),
    'arabic': Font(name='Traditional Arabic', size=12, color='000000'),
    'arabic_bold': Font(name='Traditional Arabic', size=12, bold=True, color='000000')
}

# تنسيقات التعبئة - Fill Patterns
FILL_PATTERNS = {
    'header': PatternFill(start_color=COLORS['header_bg'], end_color=COLORS['header_bg'], fill_type='solid'),
    'input': PatternFill(start_color=COLORS['input_bg'], end_color=COLORS['input_bg'], fill_type='solid'),
    'calc': PatternFill(start_color=COLORS['calc_bg'], end_color=COLORS['calc_bg'], fill_type='solid'),
    'alt_row': PatternFill(start_color=COLORS['alt_row'], end_color=COLORS['alt_row'], fill_type='solid'),
    'white': PatternFill(start_color=COLORS['white'], end_color=COLORS['white'], fill_type='solid')
}

# محاذاة - Alignments
ALIGNMENTS = {
    'center': Alignment(horizontal='center', vertical='center', wrap_text=True),
    'right': Alignment(horizontal='right', vertical='center', wrap_text=True),
    'left': Alignment(horizontal='left', vertical='center', wrap_text=True),
    'top_right': Alignment(horizontal='right', vertical='top', wrap_text=True)
}

# حدود - Borders
def create_border(style='thin', color='000000'):
    return Border(
        left=Side(style=style, color=color),
        right=Side(style=style, color=color),
        top=Side(style=style, color=color),
        bottom=Side(style=style, color=color)
    )

BORDERS = {
    'thin': create_border('thin'),
    'medium': create_border('medium'),
    'thick': create_border('thick')
}

# إعدادات Excel
MAX_ROWS = 10000
MAX_COLUMNS = 26
DEFAULT_SHEET_TAB_COLOR = COLORS['primary']