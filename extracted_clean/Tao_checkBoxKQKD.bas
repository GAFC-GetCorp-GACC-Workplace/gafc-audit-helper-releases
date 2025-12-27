Attribute VB_Name = "Tao_checkBoxKQKD"
Option Explicit
Public Sub Check_box_kqkd(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    If Not ConfirmActiveSheetRisk("Du lieu se bi ghi de o cot J:N va an cot Z.") Then Exit Sub
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ' Cot bat dau va cac cot su dung
    Dim colJ As Long: colJ = 10   ' J = STT
    Dim colK As Long: colK = 11   ' K = CONG VIEC
    Dim colL As Long: colL = 12   ' L = CHECKBOX
    Dim colM As Long: colM = 13   ' M = YES/NO (IF cong thuc)
    Dim colN As Long: colN = 14   ' N = NOTE (moi them)
    Dim colZ As Long: colZ = 26   ' Z = LinkedCell (TRUE/FALSE, se an)
    Dim rowStart As Long: rowStart = 1
    ' Danh sach cong viec
    Dim checklist As Variant
    checklist = Array( _
        "Doanh thu kh" & ChrW(7899) & "p báo cáo thu" & ChrW(7871) & " hay không", _
        "Doanh thu có kh" & ChrW(7899) & "p báo cáo theo s" & ChrW(7843) & "n l" & ChrW(432) & ChrW(7907) & "ng (n" & ChrW(7871) & "u có)", _
        "Kh" & ChrW(7899) & "p doanh thu và giá v" & ChrW(7889) & "n c" & ChrW(7911) & "a t" & ChrW(7915) & "ng công trình hay không (n" & ChrW(7871) & "u có H" & ChrW(272) & "XD)", _
        "Ki" & ChrW(7875) & "m tra quy trình ghi nh" & ChrW(7853) & "n doanh thu: h" & ChrW(7907) & "p " & ChrW(273) & ChrW(7891) & "ng, hóa " & ChrW(273) & "on, biên b" & ChrW(7843) & "n nghi" & ChrW(7879) & "m thu, giao hàng", _
        ChrW(272) & ChrW(7889) & "i v" & ChrW(7899) & "i ho" & ChrW(7841) & "t " & ChrW(273) & ChrW(7897) & "ng xu" & ChrW(7845) & "t kh" & ChrW(7849) & "u có t" & ChrW(7901) & " khai h" & ChrW(7843) & "i quan, h" & ChrW(7907) & "p " & ChrW(273) & ChrW(7891) & "ng, commercial invoice, Bill of lading, xx", _
        "Ki" & ChrW(7875) & "m tra ngày trên t" & ChrW(7901) & " khai h" & ChrW(7843) & "i quan ph" & ChrW(7847) & "n " & ChrW(273) & "ã hoàn t" & ChrW(7845) & "t", _
        "Ki" & ChrW(7875) & "m tra cut off", _
        "Thu th" & ChrW(7853) & "p b" & ChrW(7897) & " ch" & ChrW(7913) & "ng t" & ChrW(7915), _
        "Phân tích doanh thu và l" & ChrW(227) & "i g" & ChrW(7897) & "p và gi" & ChrW(7843) & "i thích vì sao có bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng (Note " & ChrW(7903) & " giá v" & ChrW(7889) & "n nam nay)", _
        "Có ki" & ChrW(7875) & "m tra b" & ChrW(7843) & "ng tính giá thành? Ph" & ChrW(432) & ChrW(417) & "ng pháp tính là gì (Note " & ChrW(7903) & " giá v" & ChrW(7889) & "n nam nay)", _
        "Giá v" & ChrW(7889) & "n có kh" & ChrW(7899) & "p v" & ChrW(7899) & "i b" & ChrW(7843) & "ng nh" & ChrW(7853) & "p xu" & ChrW(7845) & "t t" & ChrW(7891) & "n thành ph" & ChrW(7849) & "m / hàng hóa?", _
        "Ki" & ChrW(7875) & "m tra doanh thu và giá v" & ChrW(7889) & "n có matching không", _
        ChrW(272) & "ã tính toán l" & ChrW(7841) & "i giá v" & ChrW(7889) & "n ch" & ChrW(432) & "a", _
        "Phân tích bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng doanh thu ch" & ChrW(432) & "a", _
        "Phân tích bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng giá v" & ChrW(7889) & "n ch" & ChrW(432) & "a", _
"Phân tích bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng chi phí bán hàng ch" & ChrW(432) & "a", _
        "Phân tích bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng chi phí qu" & ChrW(7843) & "n lý ch" & ChrW(432) & "a" _
    )
    Dim n As Long: n = UBound(checklist) + 1
    Dim lastRow As Long: lastRow = rowStart + n
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' Xoa checkbox cu trong cot L trong pham vi bang
    Dim cb As CheckBox
    For Each cb In ws.CheckBoxes
        If Not Intersect(cb.TopLeftCell, ws.Columns(colL)) Is Nothing Then
            If cb.TopLeftCell.Row >= (rowStart + 1) And cb.TopLeftCell.Row <= lastRow Then
                cb.Delete
            End If
        End If
    Next cb
    ' Xoa noi dung vung J:N va cot Z (linked)
    ws.Range(ws.Cells(rowStart, colJ), ws.Cells(lastRow, colN)).Clear
    ws.Range(ws.Cells(rowStart, colZ), ws.Cells(lastRow, colZ)).Clear
    ' Tieu de
    With ws
        .Cells(rowStart, colJ).Value = "STT"
        .Cells(rowStart, colK).Value = "CÔNG VI" & ChrW(7878) & "C"
        .Cells(rowStart, colL).Value = "CHECK"
        .Cells(rowStart, colM).Value = "HOÀN THÀNH"
        .Cells(rowStart, colN).Value = "NOTE"
        .Range(.Cells(rowStart, colJ), .Cells(rowStart, colN)).Font.Bold = True
        .Range(.Cells(rowStart, colJ), .Cells(rowStart, colN)).HorizontalAlignment = xlCenter
    End With
    ' Ghi STT, CONG VIEC, chen checkbox o L, cong thuc IF o M
    Dim i As Long, r As Long
    For i = 0 To n - 1
        r = rowStart + 1 + i
        ' STT + cong viec
        ws.Cells(r, colJ).Value = i + 1
        ws.Cells(r, colK).Value = checklist(i)
        ' Mac dinh LinkedCell = False (hoac trong) -> M se hien NO
        ws.Cells(r, colZ).ClearContents
        ' Cong thuc IF tai M: =IF(Zr, "YES","NO")
        ws.Cells(r, colM).Formula = "=IF(" & ws.Cells(r, colZ).Address(False, False) & ", ""YES"", ""NO"")"
        ' Tao checkbox tai L, lien ket Zr
        Dim newCB As CheckBox
        Set newCB = ws.CheckBoxes.Add( _
            Top:=ws.Cells(r, colL).Top + 1, _
            Left:=ws.Cells(r, colL).Left + 2, _
            Width:=12, Height:=12)
        With newCB
            .Caption = ""                                      ' khong hien TRUE/FALSE
            .LinkedCell = ws.Cells(r, colZ).Address(False, False) ' TRUE/FALSE o Z (an)
            .Placement = xlMoveAndSize
        End With
    Next i
    ' Them dong nhac nho phia duoi cung (khong phai checklist)
    Dim reminderRow As Long: reminderRow = lastRow + 2
    With ws.Cells(reminderRow, colK)
        .Value = "L" & ChrW(432) & "u ý: Các ph" & ChrW(7847) & "n c" & ChrW(7847) & "n phân tích có th" & ChrW(7875) & " ghi chi ti" & ChrW(7871) & "t vào c" & ChrW(7897) & "t Note k" & ChrW(7871) & " bên, ho" & ChrW(7863) & "c ghi chú chi ti" & ChrW(7871) & "t vào ph" & ChrW(7847) & "n TMBCTC"
        .Font.Italic = True
        .Font.Color = RGB(0, 112, 192)
.WrapText = True
    End With
    ' An cot Z chua TRUE/FALSE
    ws.Columns(colZ).Hidden = True
    ' Dinh dang nho
    ws.Columns(colJ).AutoFit
    ws.Columns(colK).ColumnWidth = 81      ' CONG VIEC = 81
    ws.Columns(colL).ColumnWidth = 5
    ws.Columns(colM).ColumnWidth = 13      ' HOAN THANH = 13
    ws.Columns(colN).ColumnWidth = 30
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    InfoToast "Done"
End Sub

