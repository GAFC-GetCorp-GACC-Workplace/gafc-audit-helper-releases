Attribute VB_Name = "RWPricing"
Option Explicit
Public Sub rwtopricing(control As IRibbonControl)
    Dim wsRW As Worksheet, wsTarget As Worksheet
    Dim wb As Workbook
    Dim lastRowRW As Long, lastRowTarget As Long
    Dim i As Long, j As Long
    Dim codeTarget As String, codeRW As String
    Dim found As Boolean
    Dim resultRow As Long
    Set wb = ActiveWorkbook  ' ?? File dang m?
    Set wsRW = wb.Worksheets("R W")
    Set wsTarget = wb.Worksheets("D550.1 Pricing Testing RW-M")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    lastRowRW = wsRW.Cells(wsRW.Rows.Count, "C").End(xlUp).Row
    ' ?? Xóa c?t ghi chú cu (M)
    wsTarget.Range("M3:M" & lastRowTarget).ClearContents
    For i = 3 To lastRowTarget
        codeTarget = Trim(wsTarget.Cells(i, 2).Value)
        found = False
        If codeTarget <> "" Then
            For j = 3 To lastRowRW
                codeRW = Trim(wsRW.Cells(j, 3).Value)
                If codeRW = codeTarget Then
                    ' Ði?n d? li?u t? RW
                    wsTarget.Cells(i, 3).Value = wsRW.Cells(j, 5).Value  ' Tên hàng
                    wsTarget.Cells(i, 4).Value = wsRW.Cells(j, 15).Value ' S? ti?n (O)
                    wsTarget.Cells(i, 5).Value = wsRW.Cells(j, 14).Value ' S? lu?ng (N)
                    wsTarget.Cells(i, 6).Value = wsRW.Cells(j, 7).Value  ' ÐVT (G)
                    wsTarget.Cells(i, 7).FormulaR1C1 = "=RC[-3]/RC[-2]"  ' Ðon giá = D/E
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                wsTarget.Cells(i, 13).Value = "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y mã"
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
    MsgBox "Done", vbInformation
End Sub
