Option Explicit
Public Sub Tao_for_NKC(control As IRibbonControl)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim header As Variant
    Dim i As Integer
    ' T?o m?t workbook m?i
    Set wb = Workbooks.Add
    ' Xóa các sheet m?c d?nh (n?u c?n)
    Application.DisplayAlerts = False
    Do While wb.Sheets.Count > 1
        wb.Sheets(2).Delete
    Loop
    Application.DisplayAlerts = True
    ' Ð?t tên sheet d?u tiên là "NKC"
    Set ws = wb.Sheets(1)
    ws.Name = "NKC"
    ws.Activate
    ' Khai báo danh sách tiêu d?
    header = Array("Ng" & ChrW(224) & "y h" & ChrW(7841) & "ch to" & ChrW(225) & "n", "Ng" & ChrW(224) & "y ch" & ChrW(7913) & "ng t" & ChrW(7915), "Th" & ChrW(225) & "ng", _
                   "S" & ChrW(7889) & " h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n", "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i", "N" & ChrW(7907), "Có", "N" & ChrW(7907) & " TK", "Có TK", _
                   "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", "M" & ChrW(227) & " " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng", "T" & ChrW(234) & "n " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng", "Code")
    ' Ghi tiêu d? vào hàng 2
    For i = LBound(header) To UBound(header)
        ws.Cells(2, i + 1).Value = header(i)
    Next i
    ' Ð?nh d?ng tiêu d?
    With ws.Range("A2:M2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    ' B?t AutoFilter cho hàng 2
    ws.Range("A2:M2").AutoFilter
    ' T? d?ng di?u ch?nh d? r?ng c?t
    ws.Columns("A:M").AutoFit
    ' Kích ho?t workbook m?i
    wb.Activate
    MsgBox "Done", vbInformation
End Sub