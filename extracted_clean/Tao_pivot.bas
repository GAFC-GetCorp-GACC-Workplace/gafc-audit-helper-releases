Option Explicit
Sub Tao_Pivot_AnToan()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsPV As Worksheet, wsPVCT As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim hdrN_No As String, hdrN_Co As String, hdrN_Thang As String
    Dim hdrN_NoTK As String, hdrN_CoTK As String, hdrN_PhatSinh As String
    Dim rngHeaders As Range
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsData = wb.Worksheets("NKC")
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y sheet 'NKC'.", vbExclamation
        Exit Sub
    End If
    ' Xác d?nh vùng d? li?u: t? A2 t?i ô cu?i cùng có d? li?u (b?t k? c?t)
    lastRow = wsData.Cells.Find(What:="*", LookIn:=xlValues, _
                SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = wsData.Cells.Find(What:="*", LookIn:=xlValues, _
                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If lastRow < 2 Then
        MsgBox "Kh" & ChrW(244) & "ng c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u (" & ChrW(237) & "t nh" & ChrW(7845) & "t ph" & ChrW(7843) & "i c" & ChrW(243) & " header " & ChrW(7903) & " d" & ChrW(242) & "ng 2).", vbExclamation
        Exit Sub
    End If
    Set dataRange = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, lastCol))
    ' L?y tên header t? hàng 2 (d?m b?o dùng dúng chính t? trên sheet)
    Set rngHeaders = wsData.Rows(2)
    ' N?u c?t c? d?nh v? trí (A..M), có th? d?c theo index:
    hdrN_Thang = Trim(CStr(wsData.Cells(2, 3).Value))        ' c?t C
    hdrN_No = Trim(CStr(wsData.Cells(2, 6).Value))           ' c?t F (N?)
    hdrN_Co = Trim(CStr(wsData.Cells(2, 7).Value))           ' c?t G (Có)
    hdrN_NoTK = Trim(CStr(wsData.Cells(2, 8).Value))         ' c?t H (N? TK)
    hdrN_CoTK = Trim(CStr(wsData.Cells(2, 9).Value))         ' c?t I (Có TK)
    hdrN_PhatSinh = Trim(CStr(wsData.Cells(2, 10).Value))    ' c?t J (Phát sinh N?)
    ' Ki?m tra t?n t?i header thi?t y?u
    If hdrN_Thang = "" Or hdrN_No = "" Or hdrN_Co = "" Or hdrN_PhatSinh = "" Then
        MsgBox "M" & ChrW(7897) & "t ho" & ChrW(7863) & "c nhi" & ChrW(7873) & "u ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & " c" & ChrW(7847) & "n thi" & ChrW(7871) & "t b" & ChrW(7883) & " tr" & ChrW(7889) & "ng. Ki" & ChrW(7875) & "m tra h" & ChrW(224) & "ng ti" & ChrW(234) & "u " & ChrW(273) & ChrW(7873) & " (d" & ChrW(242) & "ng 2)."", vbExclamation"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' Xóa sheet cu n?u t?n t?i
    On Error Resume Next
    wb.Worksheets("PV").Delete
    wb.Worksheets("PVCT").Delete
    On Error GoTo 0
    ' T?o sheet PV
    Set wsPV = wb.Worksheets.Add(After:=wsData)
    wsPV.Name = "PV"
    ' T?o PivotCache t? Range object (an toàn)
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    ' Pivot 1 trên PV (? A4)
    Set pt = wsPV.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPV.Range("A4"), TableName:="PT_PV_1")
    With pt
        ' S? d?ng tên header l?y t? sheet d? tránh l?i do ký t?
        .PivotFields(hdrN_No).Orientation = xlPageField
        .PivotFields(hdrN_Thang).Orientation = xlRowField
        .PivotFields(hdrN_Co).Orientation = xlColumnField
        .AddDataField .PivotFields(hdrN_PhatSinh), "T" & ChrW(7893) & "ng ti" & ChrW(7873) & "n", xlSum
        .DataBodyRange.NumberFormat = "#,##0"
    End With
    ' Pivot 2 trên PV (? A26) - d?i filter/column
    Set pt = wsPV.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPV.Range("A26"), TableName:="PT_PV_2")
    With pt
        .PivotFields(hdrN_Co).Orientation = xlPageField
        .PivotFields(hdrN_Thang).Orientation = xlRowField
        .PivotFields(hdrN_No).Orientation = xlColumnField
        .AddDataField .PivotFields(hdrN_PhatSinh), "T" & ChrW(7893) & "ng ti" & ChrW(7873) & "n", xlSum
         .DataBodyRange.NumberFormat = "#,##0"
    End With
    ' T?o sheet PVCT
    Set wsPVCT = wb.Worksheets.Add(After:=wsPV)
    wsPVCT.Name = "PVCT"
    ' Pivot 1 trên PVCT (? A4): N? TK filter, Tháng rows, Có TK columns
    Set pt = wsPVCT.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPVCT.Range("A4"), TableName:="PT_PVCT_1")
    With pt
        .PivotFields(hdrN_NoTK).Orientation = xlPageField
        .PivotFields(hdrN_Thang).Orientation = xlRowField
        .PivotFields(hdrN_CoTK).Orientation = xlColumnField
        .AddDataField .PivotFields(hdrN_PhatSinh), "T" & ChrW(7893) & "ng ti" & ChrW(7873) & "n", xlSum
         .DataBodyRange.NumberFormat = "#,##0"
    End With
    ' Pivot 2 trên PVCT (? A26): Có TK filter, Tháng rows, N? TK columns
    Set pt = wsPVCT.PivotTables.Add(PivotCache:=pc, TableDestination:=wsPVCT.Range("A26"), TableName:="PT_PVCT_2")
    With pt
        .PivotFields(hdrN_CoTK).Orientation = xlPageField
        .PivotFields(hdrN_Thang).Orientation = xlRowField
        .PivotFields(hdrN_NoTK).Orientation = xlColumnField
        .AddDataField .PivotFields(hdrN_PhatSinh), "T" & ChrW(7893) & "ng ti" & ChrW(7873) & "n", xlSum
         .DataBodyRange.NumberFormat = "#,##0"
    End With
    ' M?t s? format co b?n
    wsPV.Cells.ColumnWidth = 14
    wsPVCT.Cells.ColumnWidth = 14
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Done", vbInformation
End Sub