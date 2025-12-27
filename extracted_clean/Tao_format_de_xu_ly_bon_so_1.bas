Attribute VB_Name = "Tao_format_de_xu_ly_bon_so_1"
Option Explicit
Public Sub TaoTB(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim ws As Worksheet
    Dim wsOld As Worksheet
    Dim wb As Workbook
    Dim wsExisting As Worksheet
    Dim prevAlerts As Boolean
    Set wb = ActiveWorkbook ' ?? Làm vi?c v?i file dang m?
    If wb Is Nothing Then Exit Sub
    Set wsExisting = GetSheet(wb, "Xu_ly")
    If Not wsExisting Is Nothing Then
        If Not ConfirmProceed("Sheet 'Xu_ly' da ton tai. Xoa va tao lai? Du lieu se bi mat.") Then Exit Sub
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        wsExisting.Delete
        Application.DisplayAlerts = prevAlerts
    End If
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
    InfoToast "Done"
End Sub
