Attribute VB_Name = "ChinhformatNKC"
Option Explicit
Sub Chinh_Format_NKC()
    Dim wb As Workbook
        Dim ws As Worksheet
    Dim wsNkc As Worksheet
    Dim lastRow As Long
    ' L?y sheet dang active (ví d?: NKC)
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If ActiveSheet Is Nothing Then Exit Sub
    If StrComp(ActiveSheet.Name, "NKC", vbTextCompare) <> 0 Then
        Set wsNkc = GetSheet(wb, "NKC")
        If wsNkc Is Nothing Then
            MsgBox "Khong tim thay sheet 'NKC'.", vbExclamation
            Exit Sub
        End If
        If ConfirmProceed("Sheet hien tai khong phai 'NKC'. Chuyen sang 'NKC' de format?") Then
            wsNkc.Activate
        Else
            Exit Sub
        End If
    End If
    Set ws = ActiveSheet
    ' Xác d?nh dòng cu?i cùng có d? li?u ? c?t A
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "?? Không có d? li?u t? dòng 3 tr? di!", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' ??? C?t A và B - format ngày dd/mm/yyyy
    ws.Range("A3:B" & lastRow).NumberFormat = "dd/mm/yyyy"
    ' ?? C?t C - l?y tháng c?a c?t A, paste value
    ws.Range("C3:C" & lastRow).FormulaR1C1 = "=MONTH(RC[-2])"
    ws.Range("C3:C" & lastRow).Value = ws.Range("C3:C" & lastRow).Value
    ' ?? C?t F = LEFT 3 c?a c?t H, paste value
    ws.Range("F3:F" & lastRow).FormulaR1C1 = "=LEFT(RC[2],3)"
    ws.Range("F3:F" & lastRow).Value = ws.Range("F3:F" & lastRow).Value
    ' ?? C?t G = LEFT 3 c?a c?t I, paste value
    ws.Range("G3:G" & lastRow).FormulaR1C1 = "=LEFT(RC[2],3)"
    ws.Range("G3:G" & lastRow).Value = ws.Range("G3:G" & lastRow).Value
    ' ?? C?t J - d?nh d?ng s? có d?u ph?y phân cách nghìn
    ws.Range("J3:J" & lastRow).NumberFormat = "#,##0"
    ' ?? Ô J1 - subtotal 9 (t?ng c?t J t? J3 d?n dòng cu?i)
    ws.Range("J1").Formula = "=SUBTOTAL(9,J3:J" & lastRow & ")"
    ' Format ô J1
    With ws.Range("J1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 0) ' màu vàng
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
    End With
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
