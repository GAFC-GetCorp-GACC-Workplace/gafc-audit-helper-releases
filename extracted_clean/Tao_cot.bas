Attribute VB_Name = "Tao_cot"
Option Explicit

' Show dialog to choose between Raw or Template
Public Sub Tao_NKC_Chon(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    ' Show UserForm to let user choose
    frmCreateNKC.SelectedMode = ""
    frmCreateNKC.Show

    ' Check what user selected
    If frmCreateNKC.SelectedMode = "raw" Then
        ' Create Raw workbook
        Tao_NKC Nothing
    ElseIf frmCreateNKC.SelectedMode = "template" Then
        ' Create Template workbook
        Tao_Template_NKC_TB Nothing
    End If

    ' Unload form
    Unload frmCreateNKC
End Sub

Public Sub Tao_TH(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsAfter As Worksheet
    Dim wsTH As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsAfter = wb.Worksheets("TB")
    On Error GoTo 0
    If wsAfter Is Nothing Then
        On Error Resume Next
        Set wsAfter = wb.Worksheets(wb.Worksheets.Count)
        On Error GoTo 0
    End If
    Set wsTH = Tao_TH_Template(wb, wsAfter)
End Sub
Public Sub Tao_NKC(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsNKC As Worksheet, wsTB As Worksheet
    Set wb = Workbooks.Add
    Set wsNKC = wb.Sheets(1)
    With wsNKC
        .Name = "So Nhat Ky Chung"
        .Range("A1:G1").Value = Array("M" & ChrW(227) & " CT", "Ng" & ChrW(224) & "y ho" & ChrW(7841) & "ch to" & ChrW(225) & "n", "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i", "T" & ChrW(224) & "i kho" & ChrW(7843) & "n", "PS n" & ChrW(7907), "PS c" & ChrW(243), "Kh" & ChrW(225) & "c")
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").AutoFilter
        .Columns("A:G").AutoFit
    End With
    Set wsTB = wb.Sheets.Add(After:=wsNKC)
    With wsTB
        .Name = "TB"
        .Cells(2, 5).Value = ChrW(272) & ChrW(7847) & "u k" & ChrW(7923)
        .Cells(2, 7).Value = "Ph" & ChrW(225) & "t sinh"
        .Cells(2, 9).Value = "Cu" & ChrW(7889) & "i k" & ChrW(7923)
        .Cells(2, 5).Font.Bold = True
        .Cells(2, 7).Font.Bold = True
        .Cells(2, 9).Font.Bold = True
        .Cells(3, 1).Value = "Ph" & ChrW(226) & "n c" & ChrW(244) & "ng"
        .Cells(3, 2).Value = "C" & ChrW(7845) & "p TK"
        .Cells(3, 3).Value = "TK'"
        .Cells(3, 4).Value = "T" & ChrW(234) & "n TK"
        .Cells(3, 5).Value = "N" & ChrW(7907)
        .Cells(3, 6).Value = "C" & ChrW(243)
        .Cells(3, 7).Value = "N" & ChrW(7907)
        .Cells(3, 8).Value = "C" & ChrW(243)
        .Cells(3, 9).Value = "N" & ChrW(7907)
        .Cells(3, 10).Value = "C" & ChrW(243)
        .Range("A3:J3").Font.Bold = True
        .Range("A3:J3").Interior.Color = RGB(220, 230, 241)
        .Range("A3:J3").AutoFilter
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:C").ColumnWidth = 18
        .Columns("D:D").ColumnWidth = 35
        .Columns("E:J").ColumnWidth = 15
        .Rows(4).Select
        ActiveWindow.FreezePanes = True
        .Cells(4, 1).Select
    End With
    wsNKC.Activate
    ' silent
End Sub

' Create processed template (NKC + TB) for manual data entry
Public Sub Tao_Template_NKC_TB(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsNKC As Worksheet, wsTB As Worksheet

    ' Create new workbook
    Set wb = Workbooks.Add
    Application.ScreenUpdating = False

    ' Create NKC sheet (processed format)
    Set wsNKC = wb.Sheets(1)
    With wsNKC
        .Name = "NKC"
        ' Header row
        .Cells(2, 1).Value = "Ng" & ChrW(224) & "y ho" & ChrW(7841) & "ch to" & ChrW(225) & "n"
        .Cells(2, 2).Value = "Ng" & ChrW(224) & "y ch" & ChrW(7913) & "ng t" & ChrW(7915)
        .Cells(2, 3).Value = "Th" & ChrW(225) & "ng"
        .Cells(2, 4).Value = "S" & ChrW(7889) & " h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
        .Cells(2, 5).Value = "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i"
        .Cells(2, 6).Value = "N" & ChrW(7907)
        .Cells(2, 7).Value = "C" & ChrW(243)
        .Cells(2, 8).Value = "N" & ChrW(7907) & " TK"
        .Cells(2, 9).Value = "C" & ChrW(243) & " TK"
        .Cells(2, 10).Value = "S" & ChrW(7889) & " ti" & ChrW(7873) & "n"

        ' Format header
        .Range("A2:J2").Font.Bold = True
        .Range("A2:J2").Interior.Color = RGB(220, 230, 241)
        .Range("A2:J2").AutoFilter
        .Columns("A:J").AutoFit
        ' Nút xóa lọc nhanh ở H1
        ' Xoa button cu (neu co) va tao button Xoa loc gon trong cot H
        Dim btn As Button, b As Button
        For Each b In .Buttons
            If b.Top >= .Cells(1, 8).Top - 0.5 And b.Top < .Cells(2, 8).Top Then
                b.Delete
            End If
        Next b
        Set btn = .Buttons.Add(0, 0, 10, 10)
        With btn
            .Name = "btnClearFilter_NKC"
            .Caption = "X" & ChrW(243) & "a l" & ChrW(7885) & "c"
            .OnAction = "Clear_NKC_Filter"
            .Placement = xlMoveAndSize
            .Top = .Parent.Cells(1, 8).Top + 1
            .Left = .Parent.Cells(1, 8).Left + 1
            .Width = .Parent.Cells(1, 8).Width - 2
            .Height = .Parent.Rows(1).Height - 2
            .Characters.Font.Size = 9
        End With
    End With

    ' Create TB sheet
    Set wsTB = wb.Sheets.Add(After:=wsNKC)
    With wsTB
        .Name = "TB"
        ' Period headers
        .Cells(2, 5).Value = ChrW(272) & ChrW(7847) & "u k" & ChrW(7923)
        .Cells(2, 7).Value = "Ph" & ChrW(225) & "t sinh"
        .Cells(2, 9).Value = "Cu" & ChrW(7889) & "i k" & ChrW(7923)
        .Cells(2, 5).Font.Bold = True
        .Cells(2, 7).Font.Bold = True
        .Cells(2, 9).Font.Bold = True

        ' Column headers
        .Cells(3, 1).Value = "Ph" & ChrW(226) & "n c" & ChrW(244) & "ng"
        .Cells(3, 2).Value = "C" & ChrW(7845) & "p TK"
        .Cells(3, 3).Value = "TK'"
        .Cells(3, 4).Value = "T" & ChrW(234) & "n TK"
        .Cells(3, 5).Value = "N" & ChrW(7907)
        .Cells(3, 6).Value = "C" & ChrW(243)
        .Cells(3, 7).Value = "N" & ChrW(7907)
        .Cells(3, 8).Value = "C" & ChrW(243)
        .Cells(3, 9).Value = "N" & ChrW(7907)
        .Cells(3, 10).Value = "C" & ChrW(243)

        ' Format header
        .Range("A3:J3").Font.Bold = True
        .Range("A3:J3").Interior.Color = RGB(220, 230, 241)
        .Range("A3:J3").AutoFilter

        ' Column widths
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 8
        .Columns("C:C").ColumnWidth = 18
        .Columns("D:D").ColumnWidth = 35
        .Columns("E:J").ColumnWidth = 15

        ' Freeze panes
        .Rows(4).Select
        ActiveWindow.FreezePanes = True
        .Cells(4, 1).Select
    End With

    wsNKC.Activate
    Application.ScreenUpdating = True
    
End Sub
Public Function Tao_TH_Template(wb As Workbook, afterSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("TH").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    If afterSheet Is Nothing Then
        Set ws = wb.Sheets.Add
    Else
        Set ws = wb.Sheets.Add(After:=afterSheet)
    End If
    ws.Name = "TH"
    With ws
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
        .Columns("A:A").ColumnWidth = 8
        .Columns("B:B").ColumnWidth = 11
        .Columns("C:D").ColumnWidth = 18
        .Columns("E:E").ColumnWidth = 11
        .Columns("G:H").ColumnWidth = 2
        .Columns("I:K").ColumnWidth = 16
        ' Header khu tham so (bo cuc gon, de doc)
        .Range("B1:D4").Borders.LineStyle = xlContinuous
        .Range("B1:D4").Borders.Weight = xlThin
        .Range("B1:D4").Interior.Color = RGB(255, 255, 235)
        .Range("B1:C1").Merge
        .Range("B2:C2").Merge
        .Range("B3:C3").UnMerge
        ' O nhap tien to TK (B4) - de nguoi dung de y hon
        .Range("B4").NumberFormat = "@"
        .Range("B4").Interior.Color = RGB(255, 255, 204)
        With .Range("B4").Validation
            .Delete
            .Add Type:=xlValidateInputOnly
            .IgnoreBlank = True
            .InCellDropdown = False
            .InputTitle = "Nh" & ChrW(7853) & "p TK g" & ChrW(7889) & "c"
            .InputMessage = "Nh" & ChrW(7853) & "p ti" & ChrW(7873) & "n t" & ChrW(7889) & " TK (vd 112), D3 c" & ChrW(7845) & "p TK " & ChrW(273) & "t" & ChrW(7889) & "i " & ChrW(432) & "ng."
        End With
        .Range("B1").Value = "T" & ChrW(225) & "ch s" & ChrW(7889) & " " & ChrW(226) & "m ri" & ChrW(234) & "ng"
        .Range("D1").Value = True
        .Range("C1:D1").Interior.Color = RGB(255, 240, 150)
        .Range("C1:D1").Font.Bold = True
        .Range("B2").Value = "Th" & ChrW(225) & "ng"
        .Range("B3").Value = "TK g" & ChrW(7889) & "c"
        .Range("C3").Value = "C" & ChrW(7845) & "p " & ChrW(273) & ChrW(7889) & "i " & ChrW(7913) & "ng"
        .Range("B3").Font.Bold = True
        .Range("C3").Font.Bold = True
        .Range("C2:D2").Interior.Color = RGB(255, 255, 204)
        .Range("C2:D2").Font.Bold = True
        .Range("D2").NumberFormat = "0"
        .Range("D3").Value = 4
        .Range("C2:D3").Interior.Color = RGB(255, 255, 204)
        .Range("C2:D3").Font.Bold = True
        .Range("D3").HorizontalAlignment = xlCenter
        .Range("C4").Value = "" ' Tai khoan nhap tay
        .Range("C4").Interior.Color = RGB(255, 255, 0)
        .Range("C4").Font.Bold = True
        ' Validation dropdowns
        With .Range("D1").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="TRUE,FALSE"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        With .Range("D4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="TRUE,FALSE"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        With .Range("D2").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="1,2,3,4,5,6,7,8,9,10,11,12"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Range("E4").Value = "R" & ChrW(250) & "t g" & ChrW(7885) & "n"
        .Range("E4").Font.Bold = True
        .Range("E4").Interior.Color = RGB(232, 240, 255)
        ' Default toggles
        .Range("D1").Value = True
        .Range("D4").Value = True
        ' Khu tong hop - chi tao SDDK, SPS/SDCK se duoc Auto_Tinh_TH tu dong tao
        .Range("A5").Value = "SDDK"
        .Range("A5").Font.Color = RGB(0, 0, 200)
        .Range("A5").Font.Bold = True
        ' Luu y: Khong tao SPS/SDCK o day vi vi tri dong se thay doi theo so luong TK
        ' Khu bieu do / bang phu
        .Range("H1").Value = "Bi" & ChrW(7875) & "u " & ChrW(273) & ChrW(7891)
        .Range("H1:I1").Merge
        .Range("H1:I1").Interior.Color = RGB(255, 255, 0)
        .Range("H1:I1").Font.Bold = True
        .Range("H1:I1").HorizontalAlignment = xlCenter
        .Range("K1").Value = True
        .Range("K1").Interior.Color = RGB(255, 255, 204)
        .Range("I4").Value = "T" & ChrW(224) & "i kho" & ChrW(7843) & "n"
        .Range("J4").Value = ""
        .Range("I5").Value = ChrW(272) & ChrW(7889) & "i " & ChrW(7913) & "ng"
        .Range("J5").Value = "N/C"
        With .Range("J5").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="N-C,N/C"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        .Range("I6").Value = "Ph" & ChrW(225) & "t sinh n" & ChrW(7907)
        .Range("K6").Value = "Ph" & ChrW(225) & "t sinh c" & ChrW(243)
        .Range("I6:K6").Font.Bold = True
        .Range("I6:K6").Interior.Color = RGB(0, 0, 0)
        .Range("I6:K6").Font.Color = RGB(255, 255, 255)
        For r = 7 To 18
            .Cells(r, 10).Value = Format$(r - 6, "00") ' cot J danh so dong
        Next r
        .Range("J7:J18").Font.Bold = True
        .Range("J7:J18").HorizontalAlignment = xlCenter
        ' Mau bang phat sinh
        .Range("I7:K18").Interior.Color = RGB(210, 239, 252)
        .Range("I7:K18").Borders.Color = RGB(180, 180, 180)
        ' Data bars de de xem cmax (nhu heatmap nhe)
        With .Range("I7:I18").FormatConditions.AddDatabar
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
            .BarColor.Color = RGB(255, 96, 96)
        End With
        With .Range("K7:K18").FormatConditions.AddDatabar
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
            .BarColor.Color = RGB(96, 176, 255)
        End With
        .Range("I19").Formula = "=SUM(I7:I18)"
        .Range("K19").Formula = "=SUM(K7:K18)"
        .Range("I19:K19").Font.Bold = True
        .Range("I19:K19").Interior.Color = RGB(0, 0, 0)
        .Range("I19:K19").Font.Color = RGB(255, 255, 255)
        ' Format so
        .Range("C:D").NumberFormat = "#,##0"
        .Range("I:K").NumberFormat = "#,##0"
        .Range("C5:D13").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B5:E5").Borders(xlEdgeBottom).Weight = xlThin
        .Range("B12:E12").Borders(xlEdgeTop).Weight = xlThin
    End With
    Set Tao_TH_Template = ws
End Function
