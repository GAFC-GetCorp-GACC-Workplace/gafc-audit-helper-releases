Attribute VB_Name = "FormatRW"
Option Explicit
Public Sub taorw(control As IRibbonControl)
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("R W").Delete ' Xóa n?u có s?n
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = "R W"
    ' G?p dòng tiêu d?
    ws.Range("A1:G1").Merge
    ws.Range("A1").Value = "Th" & ChrW(244) & "ng tin s" & ChrW(7843) & "n ph" & ChrW(7849) & "m"
    ws.Range("H1:I1").Merge
    ws.Range("H1").Value = ChrW(272) & ChrW(7847) & "u k" & ChrW(236)
    ws.Range("J1:K1").Merge
    ws.Range("J1").Value = "Nh" & ChrW(7853) & "p kho"
    ws.Range("L1:M1").Merge
    ws.Range("L1").Value = "Xu" & ChrW(7845) & "t kho"
    ws.Range("N1:O1").Merge
    ws.Range("N1").Value = "H" & ChrW(224) & "ng t" & ChrW(7891) & "n kho"
    ' Dòng 2
    Dim headers As Variant
    headers = Array( _
    "T" & ChrW(234) & "n kho", _
    "M" & ChrW(227) & " t" & ChrW(224) & "i kho" & ChrW(7843) & "n", _
    "M" & ChrW(227) & " s" & ChrW(7843) & "n ph" & ChrW(7849) & "m", _
    "M" & ChrW(227) & " ghi ch" & ChrW(250), _
    "T" & ChrW(234) & "n h" & ChrW(224) & "ng", _
    "T" & ChrW(234) & "n h" & ChrW(224) & "ng (Ti" & ChrW(7871) & "ng Anh)", _
    ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & " t" & ChrW(237) & "nh", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", _
    "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", _
    "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", _
    "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", _
    "S" & ChrW(7889) & " ti" & ChrW(7873) & "n" _
)
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i
    ' Can gi?a và in d?m
    With ws.Range("A1:O2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ' Tô màu tab sheet là vàng nh?t
    ws.Tab.Color = RGB(255, 255, 153) ' Vàng nh?t
    ' T? d?ng ch?nh d? r?ng
    ws.Columns("A:O").AutoFit
    ' Thêm filter cho dòng 2
    ws.Range("A2:O2").AutoFilter
    MsgBox "Done", vbInformation
End Sub

