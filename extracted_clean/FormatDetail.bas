Attribute VB_Name = "FormatDetail"
Option Explicit
Public Sub taodetail(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("D550.1.1 Detail Input").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = "D550.1.1 Detail Input"
    ' G?p dòng 1
    ws.Range("A1:I1").Merge
    ws.Range("A1").Value = "Ch" & ChrW(7913) & "ng t" & ChrW(7915)
    ws.Range("J1:M1").Merge
    ws.Range("J1").Value = "Th" & ChrW(244) & "ng tin ch" & ChrW(7913) & "ng t" & ChrW(7915)
    ws.Range("N1:O1").Merge
    ws.Range("N1").Value = "T" & ChrW(7891) & "n " & ChrW(273) & ChrW(7847) & "u k" & ChrW(236)
    ws.Range("P1:Q1").Merge
    ws.Range("P1").Value = "Nh" & ChrW(7853) & "p kho"
    ws.Range("R1:S1").Merge
    ws.Range("R1").Value = "Xu" & ChrW(7845) & "t kho"
    ws.Range("T1:U1").Merge
    ws.Range("T1").Value = "H" & ChrW(224) & "ng t" & ChrW(7891) & "n kho"
    ws.Range("V1").Value = "Gi" & ChrW(225) & " trung b" & ChrW(236) & "nh"
    ' Dòng 2
    Dim headers As Variant
headers = Array( _
    "S" & ChrW(7889) & " ch" & ChrW(7913) & "ng t" & ChrW(7915), _
    "T" & ChrW(234) & "n kh" & ChrW(225) & "ch h" & ChrW(224) & "ng", _
    "Ng" & ChrW(224) & "y ph" & ChrW(225) & "t h" & ChrW(224) & "nh", _
    "M" & ChrW(227) & " s" & ChrW(7843) & "n ph" & ChrW(7849) & "m", _
    "T" & ChrW(234) & "n h" & ChrW(224) & "ng", _
    "T" & ChrW(234) & "n h" & ChrW(224) & "ng (Ti" & ChrW(7871) & "ng Anh)", _
    ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & " t" & ChrW(237) & "nh", _
    "M" & ChrW(244) & " t" & ChrW(7843) & " 1", _
    "M" & ChrW(244) & " t" & ChrW(7843) & " 2", _
    "M" & ChrW(227) & " t" & ChrW(224) & "i kho" & ChrW(7843) & "n", _
    "T" & ChrW(234) & "n t" & ChrW(224) & "i kho" & ChrW(7843) & "n", _
    "T" & ChrW(224) & "i kho" & ChrW(7843) & "n " & ChrW(273) & ChrW(7889) & "i " & ChrW(7913) & "ng", _
    ChrW(272) & ChrW(417) & "n gi" & ChrW(225), _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng", "S" & ChrW(7889) & " ti" & ChrW(7873) & "n", _
    "Gi" & ChrW(225) & " trung b" & ChrW(236) & "nh" _
)
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(2, i + 1).Value = headers(i)
    Next i
    ' Can gi?a + in d?m + wrap
    With ws.Range("A1:V2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ' Filter
    ws.Range("A2:V2").AutoFilter
    ' T? d?ng co giãn c?t
    ws.Columns("A:V").AutoFit
    ' Tô màu tab sheet = xanh nu?c bi?n nh?t
    ws.Tab.Color = RGB(173, 216, 230) ' LightBlue
    With ws
    .Columns("A").Interior.Color = RGB(255, 230, 200) ' S? CT
    .Columns("C").Interior.Color = RGB(200, 255, 255) ' Ngày
    .Columns("D").Interior.Color = RGB(200, 230, 255) ' Mã hàng
    .Columns("P").Interior.Color = RGB(255, 255, 200) ' S? lu?ng
    .Columns("Q").Interior.Color = RGB(200, 255, 200) ' Giá tr?
End With
    InfoToast "Done"
End Sub

