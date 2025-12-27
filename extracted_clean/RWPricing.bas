Attribute VB_Name = "RWPricing"
Option Explicit
Public Sub rwtopricing(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wsRW As Worksheet, wsTarget As Worksheet
    Dim wb As Workbook
    Dim lastRowRW As Long, lastRowTarget As Long
    Dim i As Long, j As Long
    Dim dictRW As Object
    Dim rwArr As Variant
    Dim codeTarget As String, codeRW As String
    Dim found As Boolean
    Dim resultRow As Long
    Set wb = ActiveWorkbook  ' ?? File dang m?
    If wb Is Nothing Then Exit Sub
    Set wsRW = RequireSheet(wb, "R W", "Chua co sheet 'R W'. Hay chay Tao RW truoc.")
    If wsRW Is Nothing Then Exit Sub
    Set wsTarget = RequireSheet(wb, "D550.1 Pricing Testing RW-M", "Chua co sheet 'D550.1 Pricing Testing RW-M'. Hay chay Tao pricing truoc.")
    If wsTarget Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    lastRowRW = wsRW.Cells(wsRW.Rows.Count, "C").End(xlUp).Row
    ' ?? Xóa c?t ghi chú cu (M)
    wsTarget.Range("M3:M" & lastRowTarget).ClearContents
    Set dictRW = CreateObject("Scripting.Dictionary")
    If lastRowRW >= 3 Then
        rwArr = wsRW.Range("C3:O" & lastRowRW).Value
        For j = 1 To UBound(rwArr, 1)
            codeRW = Trim$(CStr(rwArr(j, 1)))
            If codeRW <> "" Then
                If Not dictRW.Exists(codeRW) Then
                    dictRW.Add codeRW, Array(rwArr(j, 3), rwArr(j, 13), rwArr(j, 12), rwArr(j, 5))
                End If
            End If
        Next j
    End If
    Dim rwVal As Variant
    For i = 3 To lastRowTarget
        codeTarget = Trim$(CStr(wsTarget.Cells(i, 2).Value))
        If codeTarget <> "" Then
            If dictRW.Exists(codeTarget) Then
                rwVal = dictRW(codeTarget)
                wsTarget.Cells(i, 3).Value = rwVal(0)
                wsTarget.Cells(i, 4).Value = rwVal(1)
                wsTarget.Cells(i, 5).Value = rwVal(2)
                wsTarget.Cells(i, 6).Value = rwVal(3)
                wsTarget.Cells(i, 7).FormulaR1C1 = "=RC[-3]/RC[-2]"
            Else
                wsTarget.Cells(i, 13).Value = "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y mA?"
            End If
        End If
    Next i
    ' ?? Format s?: S? ti?n - S? lu?ng - Ðon giá
    With wsTarget
        .Columns("D:E").NumberFormat = "#,##0" ' S? ti?n và s? lu?ng
        .Columns("G:G").NumberFormat = "#,##0" ' Ðon giá
        .Columns("A:G").AutoFit
    End With
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    InfoToast "Done"
End Sub
