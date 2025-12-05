Attribute VB_Name = "FormatPricing"
Option Explicit
Public Sub taopricing(control As IRibbonControl)
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("D550.1 Pricing Testing RW-M").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = Worksheets.Add
    ws.Name = "D550.1 Pricing Testing RW-M"
    ' Dòng 1: G?p H1:L1 ghi "Hóa don cu?i cùng"
    With ws.Range("H1:L1")
        .Merge
        .Value = "H" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n cu" & ChrW(7889) & "i c" & ChrW(249) & "ng"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(173, 216, 230) ' Màu xanh nu?c bi?n nh?t
    End With
    ' Dòng 2: tiêu d? d?y d? A2:L2
    Dim headers As Variant
    headers = Array("STT", "Mã hàng", "Tên hàng", "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "ÐVT", ChrW(272) & ChrW(417) & "n gi" & ChrW(225), _
                    "Ngày", "S" & ChrW(7889) & " ch" & ChrW(7913) & "ng t" & ChrW(7915), "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "Gi" & ChrW(225) & " tr" & ChrW(7883), ChrW(272) & ChrW(417) & "n gi" & ChrW(225))
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i
    ' Ð?nh d?ng dòng tiêu d?
    With ws.Range("A2:L2")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 250) ' Màu xanh nh?t
        .HorizontalAlignment = xlCenter
    End With
    ' T? d?ng di?u ch?nh d? r?ng c?t
    ws.Columns("A:L").AutoFit
    MsgBox "Done", vbInformation
End Sub

