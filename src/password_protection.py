"""
وحدة حماية الملفات - Password Protection Module
تشفير ملف Excel بكلمة مرور
"""

import msoffcrypto
import io
import os
import shutil
from pathlib import Path

def encrypt_excel_file(input_file, output_file, password):
    """
    تشفير ملف Excel بكلمة مرور
    
    Args:
        input_file: مسار الملف الأصلي
        output_file: مسار الملف المشفر
        password: كلمة المرور
    """
    try:
        # قراءة الملف الأصلي
        with open(input_file, 'rb') as f:
            file_content = f.read()
        
        # إنشاء كائن OfficeFile
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_content))
        
        # تشفير الملف
        with open(output_file, 'wb') as encrypted_file:
            office_file.encrypt(password, encrypted_file)
        
        print(f"✅ تم تشفير الملف بنجاح")
        return True
    
    except Exception as e:
        print(f"❌ خطأ في التشفير: {e}")
        import traceback
        traceback.print_exc()
        return False

def add_password_to_excel(file_path, password):
    """
    إضافة كلمة مرور لملف Excel
    
    Args:
        file_path: مسار الملف
        password: كلمة المرور
    """
    # إنشاء ملف مؤقت
    temp_file = str(file_path) + '.encrypted'
    
    try:
        # تشفير الملف
        if encrypt_excel_file(file_path, temp_file, password):
            # استبدال الملف الأصلي
            os.remove(file_path)
            shutil.move(temp_file, file_path)
            print(f"✅ تم حماية الملف بكلمة مرور: {password}")
            return True
        else:
            # حذف الملف المؤقت إذا فشل التشفير
            if os.path.exists(temp_file):
                os.remove(temp_file)
            return False
    
    except Exception as e:
        print(f"❌ خطأ: {e}")
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return False

def protect_with_password(file_path, password='123456'):
    """
    حماية ملف Excel بكلمة مرور (واجهة مبسطة)
    
    Args:
        file_path: مسار الملف
        password: كلمة المرور (الافتراضية: 123456)
    """
    return add_password_to_excel(file_path, password)