Attribute VB_Name = "Len_so_TB"
Option Explicit
Public Sub Len_so_bao_cao(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsTB As Worksheet, wsFS As Worksheet
    Dim lastRowTB As Long, lastRowFS As Long
    Dim i As Long, j As Long
    Dim code1 As String, code2 As String
    Dim val1 As Double, val2 As Double
    Dim arrFS As Variant
    Dim tempVal2 As Double
    Dim dictFS As Object
    Dim arrFSVal As Variant
    Dim tbData As Variant
    Dim fsRow As Long
    Dim curVal As Double
    Dim fsKey As String
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim calcMode As XlCalculation
    ' ?? L?y workbook c?a sheet dang active (tránh dính file Add-in)
    Set wb = ActiveSheet.Parent
    ' ?? Ki?m tra t?n t?i 2 sheet
    On Error Resume Next
    Set wsTB = wb.Sheets("TB")
    Set wsFS = wb.Sheets("Adjusted FS")
    On Error GoTo 0
    If wsTB Is Nothing Or wsFS Is Nothing Then
        MsgBox "? Không tìm th?y sheet 'TB' ho?c 'Adjusted FS' trong file dang m?!", vbCritical, "L?i"
        Exit Sub
    End If
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ' ?? Xác d?nh dòng cu?i trong TB
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "H").End(xlUp).Row
    lastRowFS = 250  ' ho?c có th? d?t d?ng theo D c?t n?u b?n mu?n
    ' ?? Ð?c toàn b? c?t D bên Adjusted FS
    ' ?? A??c toA?n b? c?t D bA?n Adjusted FS
    arrFS = wsFS.Range("D1:D" & lastRowFS).Value
    arrFSVal = wsFS.Range("G1:G" & lastRowFS).Value
    Set dictFS = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(arrFS, 1)
        fsKey = Trim$(CStr(arrFS(j, 1)))
        If fsKey <> "" Then
            If Not dictFS.Exists(fsKey) Then dictFS.Add fsKey, j
        End If
    Next j
    tbData = wsTB.Range("A2:I" & lastRowTB).Value
    ' ?? Duy?t t?ng dA?ng trong TB
    For i = 1 To UBound(tbData, 1)
        code1 = Trim$(CStr(tbData(i, 1)))
        code2 = Trim$(CStr(tbData(i, 2)))
        val1 = 0: val2 = 0
        If IsNumeric(tbData(i, 8)) Then val1 = tbData(i, 8)
        If IsNumeric(tbData(i, 9)) Then val2 = tbData(i, 9)
        ' ?? Tru?ng h?p d?c bi?t: code1 = code2 = 4211 ho?c 4212 ? (-val1 + val2)
        If (code1 = "4211" Or code1 = "4212") And code1 = code2 Then
            If dictFS.Exists(code1) Then
                fsRow = dictFS(code1)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + (-val1 + val2)
            End If
        Else
            ' ?? CODE1: Ghi tr?c ti?p (gi? nguyA?n d?u)
            If code1 <> "" And dictFS.Exists(code1) Then
                fsRow = dictFS(code1)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + val1
            End If
            ' ?? CODE2: M?t s? mA? s? ghi A?m
            If code2 <> "" And dictFS.Exists(code2) Then
                tempVal2 = val2
                If code2 = "2141" Or code2 = "2142" Or code2 = "2143" Or _
                   code2 = "2417" Or code2 = "139" Or code2 = "159" Then
                    tempVal2 = -val2
                End If
                fsRow = dictFS(code2)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + tempVal2
            End If
        End If
    Next i
    wsFS.Range("G1:G" & lastRowFS).Value = arrFSVal
    ' ? Thông báo hoàn t?t
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.Calculation = calcMode
    MsgBox "? Ðã chuy?n d? li?u sang Adjusted FS thành công!", vbInformation, "Hoàn t?t"
End Sub

