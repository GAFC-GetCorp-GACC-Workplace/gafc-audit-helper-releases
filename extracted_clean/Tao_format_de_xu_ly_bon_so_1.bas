Option Explicit
Public Sub TaoTB(control As IRibbonControl)
    Dim ws As Worksheet
    Dim wsOld As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' ?? Làm vi?c v?i file dang m?
    ' T?o m?i sheet "TB"
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "Xu_ly"
    ' Tiêu d?
    ws.Range("A1:G1").Value = Array("T" & ChrW(224) & "i kho" & ChrW(7843) & "n", "N" & ChrW(7907), "C" & ChrW(243), "N" & ChrW(7907), "C" & ChrW(243), "N" & ChrW(7907), "C" & ChrW(243))
    ' Filter + d?nh d?ng
    ws.Range("A1:G1").AutoFilter
    With ws.Range("A1:G1").Interior
        .Pattern = xlSolid
        .Color = RGB(150, 240, 255)
    End With
    With ws.Range("A1:G1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    MsgBox "Done", vbInformation
End Sub