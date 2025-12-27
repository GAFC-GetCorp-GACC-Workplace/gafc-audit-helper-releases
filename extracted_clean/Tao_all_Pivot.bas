Attribute VB_Name = "Tao_all_Pivot"
Option Explicit
Public Sub Tao_al_pivot(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim tStart As Double
    Dim calcMode As XlCalculation
    Dim wb As Workbook
    Dim wsNkc As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsNkc = GetSheet(wb, "NKC")
    If wsNkc Is Nothing Then
        MsgBox "Khong tim thay sheet 'NKC'. Hay tao/ xu ly NKC truoc.", vbExclamation
        Exit Sub
    End If
    wsNkc.Activate
    On Error GoTo ErrHandler
    tStart = Timer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    '=== 1?? G?I SUB Ð?NH D?NG D? LI?U ===
    Call Chinh_Format_NKC
    '=== 2?? G?I SUB T?O PIVOT ===
    Call Tao_Pivot_AnToan
    '=== 3?? T?NG K?T ===
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = calcMode
    Exit Sub
ErrHandler:
    MsgBox "Loi vui long kiem tra lai tu format NKC"
    Resume CleanUp
End Sub
