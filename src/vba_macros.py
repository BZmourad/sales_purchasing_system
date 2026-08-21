"""
ملف الماكرو - VBA Macros
تعريفات الماكرو للأزرار
"""

VBA_CODE = '''
Sub AddCustomer()
    Sheets("العملاء").Activate
    Range("A3").Select
    MsgBox "أدخل بيانات العميل الجديد في الصف التالي", vbInformation, "إضافة عميل"
End Sub

Sub SearchCustomer()
    Sheets("العملاء").Activate
    Dim searchTerm As String
    searchTerm = InputBox("أدخل اسم العميل للبحث:", "بحث عن عميل")
    If searchTerm <> "" Then
        Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart).Activate
    End If
End Sub

Sub PrintCustomers()
    Sheets("العملاء").PrintOut
End Sub

Sub AddSupplier()
    Sheets("الموردين").Activate
    Range("A3").Select
    MsgBox "أدخل بيانات المورد الجديد في الصف التالي", vbInformation, "إضافة مورد"
End Sub

Sub SearchSupplier()
    Sheets("الموردين").Activate
    Dim searchTerm As String
    searchTerm = InputBox("أدخل اسم المورد للبحث:", "بحث عن مورد")
    If searchTerm <> "" Then
        Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart).Activate
    End If
End Sub

Sub PrintSuppliers()
    Sheets("الموردين").PrintOut
End Sub

Sub AddProduct()
    Sheets("المنتجات").Activate
    Range("A3").Select
    MsgBox "أدخل بيانات المنتج الجديد في الصف التالي", vbInformation, "إضافة منتج"
End Sub

Sub SearchProduct()
    Sheets("المنتجات").Activate
    Dim searchTerm As String
    searchTerm = InputBox("أدخل اسم المنتج للبحث:", "بحث عن منتج")
    If searchTerm <> "" Then
        Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart).Activate
    End If
End Sub

Sub PrintProducts()
    Sheets("المنتجات").PrintOut
End Sub

Sub AddInvoice()
    Sheets("فواتير المبيعات").Activate
    Range("A3").Select
    MsgBox "أدخل بيانات الفاتورة الجديدة في الصف التالي", vbInformation, "إضافة فاتورة"
End Sub

Sub SearchInvoice()
    Sheets("فواتير المبيعات").Activate
    Dim searchTerm As String
    searchTerm = InputBox("أدخل رقم الفاتورة للبحث:", "بحث عن فاتورة")
    If searchTerm <> "" Then
        Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart).Activate
    End If
End Sub

Sub PrintInvoices()
    Sheets("فواتير المبيعات").PrintOut
End Sub

Sub AddPurchaseInvoice()
    Sheets("فواتير المشتريات").Activate
    Range("A3").Select
    MsgBox "أدخل بيانات فاتورة المشتريات الجديدة", vbInformation, "إضافة فاتورة مشتريات"
End Sub

Sub SearchPurchaseInvoice()
    Sheets("فواتير المشتريات").Activate
    Dim searchTerm As String
    searchTerm = InputBox("أدخل رقم الفاتورة للبحث:", "بحث عن فاتورة مشتريات")
    If searchTerm <> "" Then
        Cells.Find(What:=searchTerm, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart).Activate
    End If
End Sub

Sub PrintPurchaseInvoices()
    Sheets("فواتير المشتريات").PrintOut
End Sub

Sub GoToCustomers()
    Sheets("العملاء").Activate
End Sub

Sub GoToSuppliers()
    Sheets("الموردين").Activate
End Sub

Sub GoToProducts()
    Sheets("المنتجات").Activate
End Sub

Sub GoToSales()
    Sheets("فواتير المبيعات").Activate
End Sub

Sub GoToPurchases()
    Sheets("فواتير المشتريات").Activate
End Sub

Sub GoToReports()
    Sheets("التقارير").Activate
End Sub

Sub Generate_تقرير_المبيعات()
    MsgBox "جاري إنشاء تقرير المبيعات...", vbInformation, "تقرير المبيعات"
    ' هنا يمكن إضافة كود إنشاء التقرير
End Sub

Sub Generate_تقرير_المشتريات()
    MsgBox "جاري إنشاء تقرير المشتريات...", vbInformation, "تقرير المشتريات"
End Sub

Sub Generate_تقرير_المخزون()
    MsgBox "جاري إنشاء تقرير المخزون...", vbInformation, "تقرير المخزون"
End Sub

Sub Generate_تقرير_العملاء()
    MsgBox "جاري إنشاء تقرير العملاء...", vbInformation, "تقرير العملاء"
End Sub

Sub Generate_تقرير_الموردين()
    MsgBox "جاري إنشاء تقرير الموردين...", vbInformation, "تقرير الموردين"
End Sub

Sub Generate_تقرير_الأرباح_والخسائر()
    MsgBox "جاري إنشاء تقرير الأرباح والخسائر...", vbInformation, "تقرير الأرباح والخسائر"
End Sub

Sub Generate_تقرير_الضرائب()
    MsgBox "جاري إنشاء تقرير الضرائب...", vbInformation, "تقرير الضرائب"
End Sub

Sub Generate_تقرير_المدفوعات()
    MsgBox "جاري إنشاء تقرير المدفوعات...", vbInformation, "تقرير المدفوعات"
End Sub
'''