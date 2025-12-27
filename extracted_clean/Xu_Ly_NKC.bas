Attribute VB_Name = "Xu_Ly_NKC"
Option Explicit
Public Sub Test_Xu_ly_NKC()
    Xu_ly_NKC1111 Nothing
End Sub

' Bo sung cac cot thieu cho sheet NKC da paste
Private Sub Bo_Sung_Cot_NKC(wsNKC As Worksheet)
    Dim lastRow As Long, r As Long
    Dim ngayHT As Date, ngayCT As Date
    Dim tkNo As String, tkCo As String
    Dim tkNo3 As String, tkCo3 As String

    On Error Resume Next
    lastRow = wsNKC.Cells(wsNKC.Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0

    If lastRow < 3 Then Exit Sub

    ' Đảm bảo header đủ cột (thêm "Khac" và "Can review" nếu thiếu)
    EnsureNKCHeader wsNKC, False

    Application.ScreenUpdating = False

    ' Bo sung du lieu cho cac dong
    For r = 3 To lastRow
        If wsNKC.Cells(r, 1).Value <> "" Then
            ' Lay gia tri co san
            ngayHT = wsNKC.Cells(r, 1).Value
            tkNo = Trim(CStr(wsNKC.Cells(r, 8).Value))
            tkCo = Trim(CStr(wsNKC.Cells(r, 9).Value))
            tkNo3 = Left$(tkNo, 3)
            tkCo3 = Left$(tkCo, 3)

            ' Bo sung Ngay chung tu (cot 2) = Ngay hach toan neu thieu
            If wsNKC.Cells(r, 2).Value = "" Or Not IsDate(wsNKC.Cells(r, 2).Value) Then
                wsNKC.Cells(r, 2).Value = ngayHT
            End If
            ngayCT = wsNKC.Cells(r, 2).Value

            ' Bo sung Thang (cot 3) tu Ngay hach toan
            If wsNKC.Cells(r, 3).Value = "" Then
                wsNKC.Cells(r, 3).Value = Month(ngayHT)
            End If

            ' Cot F (No) va G (Co) luon la TK rut gon cap 3 cua cot H/I
            If tkNo <> "" Then wsNKC.Cells(r, 6).Value = tkNo3
            If tkCo <> "" Then wsNKC.Cells(r, 7).Value = tkCo3
        End If
    Next r

    ' Tong tai I1/J1 bang SUBTOTAL (khong chen dong moi)
    wsNKC.Cells(1, 9).Value = "T" & ChrW(7893) & "ng :"
    wsNKC.Cells(1, 10).Formula = "=SUBTOTAL(9,J3:J" & lastRow & ")"
    wsNKC.Cells(1, 9).Font.Bold = True
    wsNKC.Cells(1, 10).Font.Bold = True
    wsNKC.Cells(1, 10).NumberFormat = "#,##0"
    wsNKC.Cells(1, 9).Font.Size = 11
    wsNKC.Cells(1, 10).Font.Size = 11

    Application.ScreenUpdating = True

    ' Đảm bảo nút Xóa lọc luôn có trên NKC
    FixClearFilterButton wsNKC
End Sub

' Đảm bảo header NKC có đủ cột "Khac" và "Can review" (khi mở file cũ)
Private Sub EnsureNKCHeader(ws As Worksheet, Optional includeReview As Boolean = False)
    Const HDR_ROW As Long = 2
    ' Đảm bảo cột Khac tại cột K; nếu thiếu thì thêm giá trị header.
    ws.Cells(HDR_ROW, 11).Value = "Kh" & ChrW(225) & "c"

    If includeReview Then
        ' Sổ chưa xử lý: thêm cột review tại L
        ws.Cells(HDR_ROW, 12).Value = "C" & ChrW(7847) & "n review"
        With ws.Range("A2:L2")
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
            .AutoFilter
        End With
        ws.Columns("A:L").AutoFit
    Else
        ' Sổ đã xử lý: chỉ tới cột K
        With ws.Range("A2:K2")
            .Font.Bold = True
            .Interior.Color = RGB(220, 230, 241)
            .AutoFilter
        End With
        ws.Columns("A:K").AutoFit
    End If

    ' Giới hạn độ rộng cột Dien giai (E) để tránh kéo quá dài khi AutoFit
    If ws.Columns(5).ColumnWidth > 50 Then ws.Columns(5).ColumnWidth = 50
End Sub

Private Function NormalizeHeaderText(ByVal s As String) As String
    s = LCase$(Trim$(CStr(s)))
    s = Replace$(s, " ", "")
    s = Replace$(s, "á", "a")
    s = Replace$(s, "à", "a")
    s = Replace$(s, "ả", "a")
    s = Replace$(s, "ã", "a")
    s = Replace$(s, "ạ", "a")
    s = Replace$(s, "â", "a")
    s = Replace$(s, "ă", "a")
    s = Replace$(s, "ấ", "a")
    s = Replace$(s, "ầ", "a")
    s = Replace$(s, "ẩ", "a")
    s = Replace$(s, "ẫ", "a")
    s = Replace$(s, "ậ", "a")
    s = Replace$(s, "ắ", "a")
    s = Replace$(s, "ằ", "a")
    s = Replace$(s, "ẳ", "a")
    s = Replace$(s, "ẵ", "a")
    s = Replace$(s, "ặ", "a")
    s = Replace$(s, "é", "e")
    s = Replace$(s, "è", "e")
    s = Replace$(s, "ẻ", "e")
    s = Replace$(s, "ẽ", "e")
    s = Replace$(s, "ẹ", "e")
    s = Replace$(s, "ê", "e")
    s = Replace$(s, "ế", "e")
    s = Replace$(s, "ề", "e")
    s = Replace$(s, "ể", "e")
    s = Replace$(s, "�
", "e")
    s = Replace$(s, "ệ", "e")
    s = Replace$(s, "í", "i")
    s = Replace$(s, "ì", "i")
    s = Replace$(s, "ỉ", "i")
    s = Replace$(s, "ĩ", "i")
    s = Replace$(s, "ị", "i")
    s = Replace$(s, "ó", "o")
    s = Replace$(s, "ò", "o")
    s = Replace$(s, "ỏ", "o")
    s = Replace$(s, "õ", "o")
    s = Replace$(s, "ọ", "o")
    s = Replace$(s, "ô", "o")
    s = Replace$(s, "ơ", "o")
    s = Replace$(s, "ố", "o")
    s = Replace$(s, "ồ", "o")
    s = Replace$(s, "ổ", "o")
    s = Replace$(s, "ỗ", "o")
    s = Replace$(s, "ộ", "o")
    s = Replace$(s, "ớ", "o")
    s = Replace$(s, "ờ", "o")
    s = Replace$(s, "ở", "o")
    s = Replace$(s, "ỡ", "o")
    s = Replace$(s, "ợ", "o")
    s = Replace$(s, "ú", "u")
    s = Replace$(s, "ù", "u")
    s = Replace$(s, "ủ", "u")
    s = Replace$(s, "ũ", "u")
    s = Replace$(s, "ụ", "u")
    s = Replace$(s, "ư", "u")
    s = Replace$(s, "ứ", "u")
    s = Replace$(s, "ừ", "u")
    s = Replace$(s, "ử", "u")
    s = Replace$(s, "ữ", "u")
    s = Replace$(s, "ự", "u")
    s = Replace$(s, "ý", "y")
    s = Replace$(s, "ỳ", "y")
    s = Replace$(s, "ỷ", "y")
    s = Replace$(s, "ỹ", "y")
    s = Replace$(s, "ỵ", "y")
    NormalizeHeaderText = s
End Function

' Create NKC template for manual data entry
Public Sub Tao_Template_NKC(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsTemplate As Worksheet
    Dim wsExisting As Worksheet

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsExisting = GetSheet(wb, "NKC")
    If Not wsExisting Is Nothing Then
        If Not ConfirmProceed("Sheet 'NKC' da ton tai. Xoa va tao template moi? Du lieu se bi mat.") Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Delete existing NKC sheet if exists
    If Not wsExisting Is Nothing Then wsExisting.Delete

    ' Create new NKC sheet
    Set wsTemplate = wb.Worksheets.Add
    wsTemplate.Name = "NKC"

    ' Create header
    With wsTemplate
        .Cells(2, 1).Value = "Ngay hach toan"
        .Cells(2, 2).Value = "Ngay chung tu"
        .Cells(2, 3).Value = "Thang"
        .Cells(2, 4).Value = "So hoa don"
        .Cells(2, 5).Value = "Dien giai"
        .Cells(2, 6).Value = "No"
        .Cells(2, 7).Value = "Co"
        .Cells(2, 8).Value = "No TK"
        .Cells(2, 9).Value = "Co TK"
        .Cells(2, 10).Value = "So tien"
        .Cells(2, 11).Value = "Khac"
        ' Format header (không cần cột review cho sổ đã xử lý)
        .Range("A2:K2").Font.Bold = True
        .Range("A2:K2").Interior.Color = RGB(220, 230, 241)
        .Range("A2:K2").AutoFilter
        .Columns("A:K").AutoFit

        ' Add instruction
        .Cells(1, 1).Value = "Template NKC - Paste your processed data starting from row 3"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 112, 192)
    End With
    ' Đảm bảo header chuẩn (Khac) sau khi tạo mới
    EnsureNKCHeader wsTemplate, False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    InfoToast "NKC template created successfully! Paste data từ dòng 3."
    FixClearFilterButton wsTemplate
End Sub
Public Sub Clear_NKC_Filter()
    Dim ws As Worksheet
    On Error Resume Next
    If ActiveWorkbook Is Nothing Then Exit Sub
    Set ws = ActiveWorkbook.Sheets("NKC")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    ws.Rows.Hidden = False
    On Error GoTo 0
    Application.ScreenUpdating = True
    FixClearFilterButton ws
End Sub
Private Sub FixClearFilterButton(ws As Worksheet)
    Dim btn As Object, found As Boolean
    Dim leftPos As Double, topPos As Double, w As Double, h As Double

    leftPos = ws.Cells(1, 8).Left + 1
    topPos = ws.Cells(1, 8).Top + 1
    w = ws.Cells(1, 8).Width - 2
    h = ws.Rows(1).Height - 2
    If w < 10 Then w = 40
    If h < 8 Then h = 14

    On Error Resume Next
    Set btn = ws.Buttons("btnClearFilter_NKC")
    On Error GoTo 0

    If btn Is Nothing Then
        ' Thử tạo form control button
        On Error Resume Next
        Set btn = ws.Buttons.Add(leftPos, topPos, w, h)
        On Error GoTo 0
        If btn Is Nothing Then
            ' Fallback Shapes.AddFormControl nếu Buttons.Add thất bại
            On Error Resume Next
            Set btn = ws.Shapes.AddFormControl(0, leftPos, topPos, w, h) ' 0 = xlButtonControl
            On Error GoTo 0
        End If
        If btn Is Nothing Then Exit Sub
        btn.Name = "btnClearFilter_NKC"
    End If

    With btn
        On Error Resume Next
        .OnAction = "Clear_NKC_Filter"
        .Placement = xlMoveAndSize
        .Top = topPos
        .Left = leftPos
        .Width = w
        .Height = h
        .Characters.Font.Size = 9
        .Caption = "X" & ChrW(243) & "a l" & ChrW(7885) & "c"
        On Error GoTo 0
    End With
End Sub
Public Sub Chinh_Format_NKC_va_Pivot(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsNKC As Worksheet
    Dim errs As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsNKC = GetSheet(wb, "NKC")
    If wsNKC Is Nothing Then
        MsgBox "Khong tim thay sheet 'NKC'. Hay tao/ xu ly NKC truoc.", vbExclamation
        Exit Sub
    End If
    If ActiveSheet Is Nothing Or StrComp(ActiveSheet.Name, "NKC", vbTextCompare) <> 0 Then
        If Not ConfirmProceed("Macro nay nen chay tren sheet 'NKC'. Chuyen sang sheet nay khong?") Then Exit Sub
        wsNKC.Activate
    End If
    On Error Resume Next
    Application.Run "Chinh_Format_NKC"
    If Err.Number <> 0 Then
        errs = errs & "- L" & ChrW(7895) & "i khi g" & ChrW(7885) & "i Chinh_Format_NKC: " & Err.Description & vbCrLf
        Err.Clear
    End If
    Application.Run "Tao_Pivot_AnToan"
    If Err.Number <> 0 Then
        errs = errs & "- L" & ChrW(7895) & "i khi g" & ChrW(7885) & "i Tao_Pivot_AnToan: " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
    If errs = "" Then
        MsgBox ChrW(272) & ChrW(227) & " ch" & ChrW(7841) & "y xong", vbInformation
    Else
        MsgBox "Ho" & ChrW(224) & "n t" & ChrW(7845) & "t, nh" & ChrW(432) & "ng c" & ChrW(243) & " l" & ChrW(7895) & "i:" & vbCrLf & errs, vbExclamation
    End If
End Sub
' Tao mau TB va test nhanh (gom vao 1 module de import 1 lan)
' Xu ly NKC -> tao sheet NKC -> tinh TB (neu co)
Public Sub Xu_Ly_NKC_TB(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsNguon As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsNguon = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0
    If Not wsNguon Is Nothing Then
        wsNguon.Activate
    End If
    ' Buoc 1: Xu ly NKC (auto tao sheet NKC)
    Xu_ly_NKC1111 control
    ' Buoc 2: Tinh TB neu sheet TB ton tai
    If WorksheetExists("TB", wb) Then
        Tinh_Toan_TB control
    Else
        MsgBox "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y sheet TB " & ChrW(273) & ChrW(7875) & " t" & ChrW(237) & "nh to" & ChrW(225) & "n!" & vbCrLf & _
               "H" & ChrW(227) & "y t" & ChrW(7841) & "o m" & ChrW(7851) & "u TB tr" & ChrW(432) & ChrW(7899) & "c (n" & ChrW(250) & "t T" & ChrW(7841) & "o M" & ChrW(7851) & "u TB).", vbExclamation
    End If
    ' Buoc 3: Tao Pivot
    On Error Resume Next
    Tao_Pivot_AnToan
    If Err.Number <> 0 Then
        MsgBox "C" & ChrW(7843) & "nh b" & ChrW(225) & "o: Tao_Pivot_AnToan kh" & ChrW(244) & "ng ch" & ChrW(7841) & "y " & ChrW(273) & ChrW(432) & ChrW(7907) & "c." & vbCrLf & _
          "Chi ti" & ChrW(7871) & "t:" & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub
Public Sub Xu_ly_NKC1111(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wsNguon As Worksheet, wsKetQua As Worksheet
    Dim wb As Workbook
    Dim dictGroup As Object
    Dim lastRow As Long, i As Long
    Dim arrData As Variant
    Dim rowCount As Long
    Dim arrMaCT() As String, arrNgay() As Variant, arrDienGiai() As Variant
    Dim arrTK() As String, arrTK3() As String
    Dim arrNo() As Double, arrCo() As Double
    Dim arrKhac() As Variant, arrMonth() As Variant
    Dim arrKey() As String
    Dim pairCache As Object
    Dim key As Variant, r As Variant
    Dim pivotErr As String, thMsg As String
    Dim wsNKCExists As Worksheet
    Dim isTemplateNKC As Boolean
    Dim includeReview As Boolean
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    ' Mac dinh: xu ly du lieu tho -> co cot Can review
    includeReview = True

    Set wb = ActiveWorkbook
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents

    ' Check if NKC sheet already exists (user used template)
    On Error Resume Next
    Set wsNKCExists = wb.Sheets("NKC")
    On Error GoTo 0

    If Not wsNKCExists Is Nothing Then
        ' Skip only if this is a manual template NKC
        isTemplateNKC = (InStr(1, UCase$(CStr(wsNKCExists.Cells(1, 1).Value)), "TEMPLATE NKC") > 0)
        If isTemplateNKC Then
            ' NKC sheet exists (template) - skip processing, go straight to next steps
            InfoToast "Ph" & ChrW(225) & "t hi" & ChrW(7879) & "n sheet NKC " & ChrW(273) & ChrW(227) & " t" & ChrW(7891) & "n t" & ChrW(7841) & "i! " & _
                     "B" & ChrW(7887) & " qua b" & ChrW(432) & ChrW(7899) & "c x" & ChrW(7917) & " l" & ChrW(253) & ", ch" & ChrW(7841) & "y ti" & ChrW(7871) & "p c" & ChrW(225) & "c b" & ChrW(432) & ChrW(7899) & "c ti" & ChrW(7871) & "p theo..."

            ' Đảm bảo header đủ cột Khac (không thêm review cho luồng đã xử lý)
            EnsureNKCHeader wsNKCExists, False
            includeReview = False

            ' Continue with next steps (TH, Pivot, etc.)
            GoTo SkipProcessing
        Else
            InfoToast "Detected existing NKC (non-template). Rebuilding from source..."
        End If
    End If

    ' Normal flow: Process raw data
    On Error Resume Next
    Set wsNguon = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0
    If wsNguon Is Nothing Then
        If ActiveSheet Is Nothing Then
            MsgBox "Khong tim thay sheet 'So Nhat Ky Chung' va khong co sheet dang active.", vbExclamation
            Exit Sub
        End If
        If Not ConfirmProceed("Khong tim thay sheet 'So Nhat Ky Chung'. Su dung sheet hien tai '" & ActiveSheet.Name & "' lam nguon?") Then Exit Sub
        Set wsNguon = ActiveSheet
    End If
    wsNguon.Activate
    Set wb = wsNguon.Parent
    lastRow = wsNguon.Cells(wsNguon.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Kh" & ChrW(244) & "ng c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u " & ChrW(273) & ChrW(7875) & " x" & ChrW(7917) & " l" & ChrW(253) & "!", vbExclamation
        Exit Sub
    End If
    ' Doc du lieu vao array truoc khi tao sheet moi
    arrData = wsNguon.Range("A2:G" & lastRow).Value
    rowCount = UBound(arrData, 1)
    ReDim arrMaCT(1 To rowCount)
    ReDim arrNgay(1 To rowCount)
    ReDim arrDienGiai(1 To rowCount)
    ReDim arrTK(1 To rowCount)
    ReDim arrTK3(1 To rowCount)
    ReDim arrNo(1 To rowCount)
    ReDim arrCo(1 To rowCount)
    ReDim arrKhac(1 To rowCount)
    ReDim arrMonth(1 To rowCount)
    ReDim arrKey(1 To rowCount)
    For i = 1 To rowCount
        arrMaCT(i) = Trim$(CStr(arrData(i, 1)))
        arrNgay(i) = arrData(i, 2)
        arrDienGiai(i) = arrData(i, 3)
        arrTK(i) = Trim$(CStr(arrData(i, 4)))
        If IsNumeric(arrData(i, 5)) Then arrNo(i) = CDbl(arrData(i, 5)) Else arrNo(i) = 0#
        If IsNumeric(arrData(i, 6)) Then arrCo(i) = CDbl(arrData(i, 6)) Else arrCo(i) = 0#
        arrKhac(i) = arrData(i, 7)
        If arrTK(i) <> "" Then
            arrTK3(i) = Left$(arrTK(i), 3)
        Else
            arrTK3(i) = ""
        End If
        arrMonth(i) = GetMonthValue(arrNgay(i))
        If arrMaCT(i) <> "" Then
            arrKey(i) = arrMaCT(i) & "|" & Trim$(CStr(arrNgay(i)))
        End If
    Next i
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' Tao sheet ket qua moi (ten: NKC)
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("NKC").Delete
    On Error GoTo 0
    Set wsKetQua = wb.Worksheets.Add(After:=wsNguon)
    wsKetQua.Name = "NKC"
    Application.DisplayAlerts = True
    ' Tao header theo mau
    With wsKetQua
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
        .Cells(2, 11).Value = "Kh" & ChrW(225) & "c"
        .Cells(2, 12).Value = "C" & ChrW(7847) & "n review"
        .Range("A2:L2").Font.Bold = True
        .Range("A2:L2").Interior.Color = RGB(220, 230, 241)
        .Range("A2:L2").AutoFilter
        .Columns("A:L").AutoFit
    End With
    ' Chuẩn hóa header cho luồng "chưa xử lý" (giữ Khac + review)
    EnsureNKCHeader wsKetQua, True
    Set dictGroup = CreateObject("Scripting.Dictionary")
    ' Nhom du lieu theo MaCT|Ngay
    For i = 1 To rowCount
        If arrMaCT(i) <> "" Then
            key = arrKey(i)
            If Not dictGroup.Exists(key) Then dictGroup.Add key, New Collection
            dictGroup(key).Add i
        End If
    Next i
    ' ========== BUOC 1: XAC DINH NHOM "BAN" ==========
    ' Nhom "ban" = co it nhat 1 dong co CA No va Co
    Dim dictDirty As Object
    Set dictDirty = CreateObject("Scripting.Dictionary")
    Set pairCache = CreateObject("Scripting.Dictionary")
    For Each key In dictGroup.keys
        Dim isDirty As Boolean
        isDirty = False
        For Each r In dictGroup(key)
            If arrNo(r) <> 0 And arrCo(r) <> 0 Then
                isDirty = True
                Exit For
            End If
        Next r
        dictDirty.Add key, isDirty
    Next key
    ' ========== XU LY VA THU THAP OUTPUT ==========
    Dim outputArr() As Variant
    Dim colCount As Long
    colCount = IIf(includeReview, 12, 11)
    ReDim outputArr(1 To (lastRow - 1) * 10, 1 To colCount)
    Dim dongOut As Long
    dongOut = 1
    For Each key In dictGroup.keys
        Dim dsNoEntries As Collection, dsCoEntries As Collection
        Set dsNoEntries = New Collection
        Set dsCoEntries = New Collection
        For Each r In dictGroup(key)
            Dim tienNoGoc As Double, tienCoGoc As Double
            tienNoGoc = arrNo(r)
            tienCoGoc = arrCo(r)
            If tienNoGoc <> 0 Then
                dsNoEntries.Add Array(r, tienNoGoc)
            End If
            If tienCoGoc <> 0 Then
                dsCoEntries.Add Array(r, tienCoGoc)
            End If
        Next r
        If dsNoEntries.Count = 0 Or dsCoEntries.Count = 0 Then GoTo NextGroup
        Dim usedNo() As Double, usedCo() As Double
        ReDim usedNo(1 To dsNoEntries.Count)
        ReDim usedCo(1 To dsCoEntries.Count)
        Dim idxNo As Long, idxCo As Long
        Dim entryNo As Variant, entryCo As Variant
        Dim rNo As Long, rCo As Long
        Dim tienNoEntry As Double, tienCoEntry As Double
        Dim tienNo As Double, tienCo As Double
        Dim tienPhanBo As Double
        Dim absNo As Double, absCo As Double
        Dim tkNo As String, tkCo As String
        Dim khacValFast As Variant
        Dim needReview As String
        ' Lay trang thai "ban" cua nhom
        needReview = ""
        If dictDirty(key) Then needReview = "X"
        ' ========== FAST PATH: 1 NO or 1 CO -> phan bo truc tiep (giu dung gia tri am) ==========
        If dsNoEntries.Count = 1 Or dsCoEntries.Count = 1 Then
            If dsNoEntries.Count = 1 Then
                entryNo = dsNoEntries(1)
                rNo = entryNo(0)
                tkNo = arrTK(rNo)
                For idxCo = 1 To dsCoEntries.Count
                    entryCo = dsCoEntries(idxCo)
                    rCo = entryCo(0)
                    tienCoEntry = entryCo(1)
                    tkCo = arrTK(rCo)
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    khacValFast = arrKhac(rNo)
                    If Len(Trim$(khacValFast)) = 0 Then khacValFast = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    outputArr(dongOut, 3) = arrMonth(rNo)
                    outputArr(dongOut, 4) = arrMaCT(rNo)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 6) = arrTK3(rNo)  ' No (3 ky tu dau)
                    outputArr(dongOut, 7) = arrTK3(rCo)  ' Co (3 ky tu dau)
                    outputArr(dongOut, 8) = arrTK(rNo)  ' No TK (full)
                    outputArr(dongOut, 9) = arrTK(rCo)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienCoEntry     ' So tien (giu dung dau)
                    outputArr(dongOut, 11) = khacValFast     ' Khac
                    If includeReview Then
                        If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                            outputArr(dongOut, 12) = needReview
                        Else
                            outputArr(dongOut, 12) = "X"
                        End If
                    End If
                    usedNo(1) = usedNo(1) + tienCoEntry
                    usedCo(idxCo) = usedCo(idxCo) + tienCoEntry
                    dongOut = dongOut + 1
                Next idxCo
            Else
                entryCo = dsCoEntries(1)
                rCo = entryCo(0)
                tkCo = arrTK(rCo)
                For idxNo = 1 To dsNoEntries.Count
                    entryNo = dsNoEntries(idxNo)
                    rNo = entryNo(0)
                    tienNoEntry = entryNo(1)
                    tkNo = arrTK(rNo)
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    khacValFast = arrKhac(rNo)
                    If Len(Trim$(khacValFast)) = 0 Then khacValFast = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    outputArr(dongOut, 3) = arrMonth(rNo)
                    outputArr(dongOut, 4) = arrMaCT(rNo)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 6) = arrTK3(rNo)  ' No (3 ky tu dau)
                    outputArr(dongOut, 7) = arrTK3(rCo)  ' Co (3 ky tu dau)
                    outputArr(dongOut, 8) = arrTK(rNo)  ' No TK (full)
                    outputArr(dongOut, 9) = arrTK(rCo)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienNoEntry     ' So tien (giu dung dau)
                    outputArr(dongOut, 11) = khacValFast     ' Khac
                    If includeReview Then
                        If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                            outputArr(dongOut, 12) = needReview
                        Else
                            outputArr(dongOut, 12) = "X"
                        End If
                    End If
                    usedNo(idxNo) = usedNo(idxNo) + tienNoEntry
                    usedCo(1) = usedCo(1) + tienNoEntry
                    dongOut = dongOut + 1
                Next idxNo
            End If
            GoTo NextGroup
        End If
        ' ========== PASS 1: Ghep theo QUY TAC KE TOAN ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass1
            tkNo = arrTK(rNo)
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoPass1
                tkCo = arrTK(rCo)
                If Abs(tienNo - tienCo) < 0.01 And IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    ' Format mau: NgayHT, NgayCT, Thang, SoHD, DienGiai, No, Co, NoTK, CoTK, SoTien, Khac, CanReview
                    Dim khacVal As Variant
                    khacVal = arrKhac(rNo)
                    If Len(Trim$(khacVal)) = 0 Then khacVal = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    ' Thang lay theo Ngay hach toan (cot A sheet NKC = col 2 nguon)
                    outputArr(dongOut, 3) = arrMonth(rNo)
                    outputArr(dongOut, 4) = arrMaCT(rNo)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 6) = arrTK3(rNo)  ' No (3 ky tu dau TK No)
                    outputArr(dongOut, 7) = arrTK3(rCo)  ' Co (3 ky tu dau TK Co)
                    outputArr(dongOut, 8) = arrTK(rNo)  ' No TK (full)
                    outputArr(dongOut, 9) = arrTK(rCo)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienNo          ' So tien
                    outputArr(dongOut, 11) = khacVal          ' Khac (lay tu G, uu tien dong No, neu trong thi dong Co)
                    usedNo(idxNo) = usedNo(idxNo) + tienNo
                    usedCo(idxCo) = usedCo(idxCo) + tienNo
                    dongOut = dongOut + 1
                    Exit For
                End If
NextCoPass1:
            Next idxCo
NextNoPass1:
        Next idxNo
        ' ========== PASS 2: Ghep so tien khop ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass2
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoPass2
                If Abs(tienNo - tienCo) < 0.01 Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    Dim khacVal2 As Variant
                    khacVal2 = arrKhac(rNo)
                    If Len(Trim$(khacVal2)) = 0 Then khacVal2 = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    outputArr(dongOut, 3) = arrMonth(rNo)
                    outputArr(dongOut, 4) = arrMaCT(rNo)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 6) = arrTK3(rNo)  ' No (3 ky tu dau)
                    outputArr(dongOut, 7) = arrTK3(rCo)  ' Co (3 ky tu dau)
                    outputArr(dongOut, 8) = arrTK(rNo)  ' No TK (full)
                    outputArr(dongOut, 9) = arrTK(rCo)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienNo          ' So tien
                    outputArr(dongOut, 11) = khacVal2        ' Khac
                    usedNo(idxNo) = usedNo(idxNo) + tienNo
                    usedCo(idxCo) = usedCo(idxCo) + tienNo
                    dongOut = dongOut + 1
                    Exit For
                End If
NextCoPass2:
            Next idxCo
NextNoPass2:
        Next idxNo
        ' ========== PASS 3: Phan bo phan con lai (uu tien so du nho truoc, luon can review) ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass3
            Do While Abs(tienNo) >= 0.01
                Dim bestIdx As Long, bestRemain As Double, bestCo As Double
                bestIdx = 0: bestRemain = 0: bestCo = 0
                ' Chon dong Co co so du nho nhat de phan bo truoc (tranh an het vao dong lon)
                For idxCo = 1 To dsCoEntries.Count
                    entryCo = dsCoEntries(idxCo)
                    tienCoEntry = entryCo(1)
                    tienCo = tienCoEntry - usedCo(idxCo)
                    If Abs(tienCo) >= 0.01 Then
                        If bestIdx = 0 Or Abs(tienCo) < bestRemain Then
                            bestIdx = idxCo
                            bestRemain = Abs(tienCo)
                            bestCo = tienCo
                        End If
                    End If
                Next idxCo
                If bestIdx = 0 Then Exit Do
                entryCo = dsCoEntries(bestIdx)
                rCo = entryCo(0)
                absNo = Abs(tienNo)
                absCo = Abs(bestCo)
                If absNo < absCo Then
                    tienPhanBo = absNo
                Else
                    tienPhanBo = absCo
                End If
                tienPhanBo = tienPhanBo * Sgn(tienNo)
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                tkNo = arrTK(rNo)
                tkCo = arrTK(rCo)
                Dim khacVal3 As Variant
                khacVal3 = arrKhac(rNo)
                If Len(Trim$(khacVal3)) = 0 Then khacVal3 = arrKhac(rCo)
                outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay hach toan
                outputArr(dongOut, 2) = ""               ' Ngay chung tu
                outputArr(dongOut, 3) = arrMonth(rNo)
                outputArr(dongOut, 4) = arrMaCT(rNo)  ' So hoa don (MaCT)
                outputArr(dongOut, 5) = arrDienGiai(rNo)  ' Dien giai
                outputArr(dongOut, 6) = arrTK3(rNo)  ' No (3 ky tu dau)
                outputArr(dongOut, 7) = arrTK3(rCo)  ' Co (3 ky tu dau)
                outputArr(dongOut, 8) = arrTK(rNo)  ' No TK (full)
                outputArr(dongOut, 9) = arrTK(rCo)  ' Co TK (full)
                outputArr(dongOut, 10) = tienPhanBo      ' So tien
                outputArr(dongOut, 11) = khacVal3        ' Khac
                ' Nếu luồng chưa xử lý (includeReview=True) thì cột 12 là review
                If includeReview Then
                    If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                        outputArr(dongOut, 12) = needReview
                    Else
                        outputArr(dongOut, 12) = "X"
                    End If
                End If
                usedNo(idxNo) = usedNo(idxNo) + tienPhanBo
                usedCo(bestIdx) = usedCo(bestIdx) + tienPhanBo
                tienNo = tienNo - tienPhanBo
                dongOut = dongOut + 1
                If Abs(tienNo) < 0.01 Then Exit Do
            Loop
NextNoPass3:
        Next idxNo
NextGroup:
    Next key
    ' ========== GHI OUTPUT ==========
    If dongOut > 1 Then
        Dim finalOut() As Variant
        ReDim finalOut(1 To dongOut - 1, 1 To colCount)
        Dim j As Long
        For i = 1 To dongOut - 1
            For j = 1 To colCount
                finalOut(i, j) = outputArr(i, j)
            Next j
        Next i
        wsKetQua.Range("A3").Resize(dongOut - 1, colCount).Value = finalOut
        ' ========== TO VANG CAC DONG CAN REVIEW ==========
        ' Tô vàng cột review (chỉ khi includeReview=True)
        Dim rng As Range
        If includeReview Then
            For i = 3 To dongOut + 1
                If wsKetQua.Cells(i, 12).Value = "X" Then
                    Set rng = wsKetQua.Range(wsKetQua.Cells(i, 1), wsKetQua.Cells(i, 12))
                    rng.Interior.Color = RGB(255, 255, 150) ' Vang
                End If
            Next i
        End If
    End If
    ' ========== FORMAT ==========
    Dim lastRowOut As Long
    lastRowOut = dongOut + 1
    wsKetQua.Cells(1, 10).Formula = "=SUBTOTAL(9,J3:J" & lastRowOut & ")"
    wsKetQua.Cells(1, 10).Font.Bold = True
    wsKetQua.Cells(1, 9).Value = "Tong:"
    wsKetQua.Cells(1, 9).Font.Bold = True
    wsKetQua.Columns("J").NumberFormat = "#,##0"
    wsKetQua.Columns("A:B").NumberFormat = "dd/mm/yyyy"
    ' Dem so dong can review
    Dim countReview As Long
    If includeReview Then
        countReview = Application.WorksheetFunction.CountIf(wsKetQua.Columns(12), "X")
    Else
        countReview = 0
    End If
    ' Dem so nhom ban
    Dim countDirty As Long
    countDirty = 0
    For Each key In dictDirty.keys
        If dictDirty(key) Then countDirty = countDirty + 1
    Next key
    ' Keep app settings off until all steps finish for speed
    ' Hien thi ket qua NKC truoc (toast tu tat)
    InfoToast "X" & ChrW(7917) & " l" & ChrW(253) & " NKC ho" & ChrW(224) & "n th" & ChrW(224) & "nh! Output: " & (dongOut - 1) & _
              "; Nh" & ChrW(243) & "m b" & ChrW(7845) & "t to" & ChrW(224) & "n: " & countDirty & _
              IIf(includeReview, "; C" & ChrW(7847) & "n review: " & countReview, "")
    ' Sau do moi tinh TB (neu co)
    Dim tbMsg As String
    If WorksheetExists("TB", wb) Then
        tbMsg = Auto_Tinh_TB(wsKetQua)
        If tbMsg <> "" Then InfoToast "T" & ChrW(205) & "NH TO" & ChrW(193) & "N TB TH" & ChrW(192) & "NH C" & ChrW(212) & "NG! " & tbMsg
    Else
        WarnToast "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y sheet TB! T" & ChrW(7841) & "o m" & ChrW(7851) & "u TB tr" & ChrW(432) & ChrW(7899) & "c r" & ChrW(7891) & "i t" & ChrW(237) & "nh."
    End If
    ' Chay Pivot sau khi xu ly NKC/TB
    On Error Resume Next
    Tao_Pivot_AnToan
    If Err.Number <> 0 Then
        pivotErr = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If pivotErr <> "" Then
        MsgBox "C" & ChrW(7842) & "NH B" & ChrW(193) & "O: kh" & ChrW(244) & "ng t" & ChrW(7841) & "o " & ChrW(273) & ChrW(432) & ChrW(7907) & "c Pivot" & vbCrLf & _
               "Chi ti" & ChrW(7871) & "t: " & pivotErr, vbExclamation
    End If
    ' Cap nhat sheet TH neu co
    thMsg = Auto_Tinh_TH(wsKetQua)
    If thMsg <> "" Then WarnToast thMsg

SkipProcessing:
    ' Bat auto refresh TH sau khi da co NKC/TB
    On Error Resume Next
    Application.Run "Enable_TH_AutoRefresh"
    On Error GoTo 0

    ' Khi skip processing (da co NKC), van can chay cac buoc tiep theo
    Dim wsNKCSkip As Worksheet
    Dim tbMsgSkip As String, thMsgSkip As String

    ' Lay sheet NKC
    On Error Resume Next
    Set wsNKCSkip = wb.Worksheets("NKC")
    On Error GoTo 0

    ' Bo sung cac cot thieu cho NKC neu can
    If Not wsNKCSkip Is Nothing Then
        Bo_Sung_Cot_NKC wsNKCSkip
    End If

    ' Buoc 1: Tinh toan TB neu co sheet TB
    If WorksheetExists("TB", wb) And Not wsNKCSkip Is Nothing Then
        tbMsgSkip = Auto_Tinh_TB(wsNKCSkip)
        If tbMsgSkip <> "" Then
            MsgBox "T" & ChrW(205) & "NH TO" & ChrW(193) & "N TB TH" & ChrW(192) & "NH C" & ChrW(212) & "NG!" & tbMsgSkip, vbInformation
        End If
    End If

    ' Buoc 2: Tao/Cap nhat TH
    Dim wsTHCheck As Worksheet
    On Error Resume Next
    Set wsTHCheck = wb.Worksheets("TH")
    On Error GoTo 0

    If wsTHCheck Is Nothing Then
        ' Chua co TH, tao moi
        On Error Resume Next
        Application.Run "Tao_TH", Nothing
        If Err.Number <> 0 Then
            MsgBox "C" & ChrW(7843) & "nh b" & ChrW(225) & "o: Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y Tao_TH procedure!", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    Else
        ' TH da co, cap nhat
        If Not wsNKCSkip Is Nothing Then
            thMsgSkip = Auto_Tinh_TH(wsNKCSkip)
            If thMsgSkip <> "" Then WarnToast thMsgSkip
        End If
    End If

    ' Buoc 3: Tao Pivot
    On Error Resume Next
    Application.Run "Tao_Pivot_AnToan"
    If Err.Number <> 0 Then
        MsgBox "C" & ChrW(7843) & "nh b" & ChrW(225) & "o: Tao_Pivot_AnToan kh" & ChrW(244) & "ng ch" & ChrW(7841) & "y " & ChrW(273) & ChrW(432) & ChrW(7907) & "c." & vbCrLf & _
               "Chi ti" & ChrW(7871) & "t: " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0

    ' Additional steps after processing (or skipping)
    MsgBox ChrW(272) & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t!" & vbCrLf & _
           "C" & ChrW(243) & " th" & ChrW(7875) & " ch" & ChrW(7841) & "y ti" & ChrW(7871) & "p c" & ChrW(225) & "c b" & ChrW(432) & ChrW(7899) & "c kh" & ChrW(225) & "c n" & ChrW(7871) & "u c" & ChrW(7847) & "n.", vbInformation
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
End Sub
Private Function Auto_Tinh_TB(wsNKC As Worksheet) As String
    Dim wsTB As Worksheet, wsSource As Worksheet
    Dim wb As Workbook
    Dim lastRowNKC As Long, lastRowTB As Long, lastRowSource As Long
    Dim r As Long, i As Long
    Dim tkFull As String, tkTrim As String
    Dim dictSourceTK As Object
    Dim tkSource As String
    Dim useLeftMatch As Boolean
    Dim warnNonNum As String
    Dim oldCalc As XlCalculation
    On Error GoTo ErrorHandler
    oldCalc = Application.Calculation
    Set wb = wsNKC.Parent
    Set wsTB = wb.Sheets("TB")
    lastRowNKC = wsNKC.Cells(wsNKC.Rows.Count, "A").End(xlUp).Row
    If lastRowNKC < 3 Then
        Auto_Tinh_TB = vbCrLf & vbCrLf & "TB " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7853) & "p nh" & ChrW(7853) & "t c" & ChrW(244) & "ng th" & ChrW(7913) & "c B, L, M tr" & ChrW(234) & "n " & (lastRowTB - 3) & " d" & ChrW(242) & "ng."
        Exit Function
    End If
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
    If lastRowTB < 4 Then
        Auto_Tinh_TB = vbCrLf & vbCrLf & "TB " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7853) & "p nh" & ChrW(7853) & "t c" & ChrW(244) & "ng th" & ChrW(7913) & "c B, L M tr" & ChrW(234) & "n " & (lastRowTB - 3) & " d" & ChrW(242) & "ng."
        Exit Function
    End If
    ' Đảm bảo cột số E:J của TB là số (chuyển text sang số nếu có)
    warnNonNum = NormalizeTBNumberColumns(wsTB, 4, lastRowTB)
    If Len(warnNonNum) > 0 Then WarnToast warnNonNum
    ' Tao dictionary chua danh sach TK tu sheet "So Nhat Ky Chung"
    Set dictSourceTK = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsSource = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo ErrorHandler
    If Not wsSource Is Nothing Then
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
        For i = 2 To lastRowSource
            tkSource = Trim$(CStr(wsSource.Cells(i, 4).Value))
            If tkSource <> "" And Not dictSourceTK.Exists(tkSource) Then
                dictSourceTK.Add tkSource, True
            End If
        Next i
    End If
    Application.Calculation = xlCalculationManual
    ' FORMAT COT C (TAI KHOAN) THANH TEXT DE TRANH BI CONVERT SANG NUMBER KHI EDIT
    wsTB.Columns(3).NumberFormat = "@"
    For r = 4 To lastRowTB
        If wsTB.Cells(r, 3).Value <> "" Then
            tkFull = CStr(wsTB.Cells(r, 3).Value)
            tkTrim = Trim$(tkFull)
            ' FORCE GHI LAI GIA TRI DE CHUYEN HOAN TOAN SANG TEXT (tranh bi luu duoi dang number)
            wsTB.Cells(r, 3).Value = "'" & tkTrim
            ' Luu do dai TK vao cot B (VALUE, khong phai formula de tranh loi khi user edit)
            wsTB.Cells(r, 2).Value = Len(tkTrim)
            ' Xac dinh logic: neu TK >= 6 ky tu va KHONG ton tai trong So Nhat Ky Chung => dung LEFT MATCH
            useLeftMatch = (Len(tkTrim) >= 6 And Not dictSourceTK.Exists(tkTrim))
            ' Cong thuc cho cot L (Lech No)
            If Len(tkTrim) = 3 Then
                ' Case 1: TK 3 ky tu - EXACT MATCH tren cot F (No 3 ky tu)
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMIFS(" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10," & wsNKC.Name & "!R3C6:R" & lastRowNKC & "C6,RC3)-RC7"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                ' Case 2: TK 4-5 ky tu - LEFT MATCH (dung RC2 thay vi LEN(RC3) de tranh bi tinh lai khi edit)
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C8:R" & lastRowNKC & "C8,RC2)=RC3)*" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10)-RC7"
            ElseIf useLeftMatch Then
                ' Case 3a: TK >= 6 ky tu NHUNG KHONG ton tai trong SNKC => TK tong hop => LEFT MATCH (dung RC2)
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C8:R" & lastRowNKC & "C8,RC2)=RC3)*" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10)-RC7"
            Else
                ' Case 3b: TK >= 6 ky tu VA ton tai trong SNKC => TK chi tiet => EXACT MATCH
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMIFS(" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10," & wsNKC.Name & "!R3C8:R" & lastRowNKC & "C8,RC3)-RC7"
            End If
            ' Cong thuc cho cot M (Lech Co) - dung RC2 thay vi LEN(RC3) de tranh bi tinh lai khi edit
            If Len(tkTrim) = 3 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10," & wsNKC.Name & "!R3C7:R" & lastRowNKC & "C7,RC3)-RC8"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C9:R" & lastRowNKC & "C9,RC2)=RC3)*" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10)-RC8"
            ElseIf useLeftMatch Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C9:R" & lastRowNKC & "C9,RC2)=RC3)*" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10)-RC8"
            Else
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(" & wsNKC.Name & "!R3C10:R" & lastRowNKC & "C10," & wsNKC.Name & "!R3C9:R" & lastRowNKC & "C9,RC3)-RC8"
            End If
        End If
    Next r
    wsTB.Calculate
    Application.Calculation = oldCalc
    ' To mau dong theo tieu chi
    For r = 4 To lastRowTB
        If wsTB.Cells(r, 3).Value <> "" Then
            Dim tkLen As Long
            tkLen = wsTB.Cells(r, 2).Value
            ' Uu tien 1: To vang dong co lech L hoac M
            If wsTB.Cells(r, 12).Value <> 0 Or wsTB.Cells(r, 13).Value <> 0 Then
                wsTB.Range("A" & r & ":M" & r).Interior.Color = RGB(255, 255, 150)
            ' Uu tien 2: To xanh (B den J) va BOLD cho dong co cap TK = 3
            ElseIf tkLen = 3 Then
                wsTB.Range("B" & r & ":J" & r).Interior.Color = RGB(146, 208, 80)  ' Xanh la nhat
                wsTB.Range("A" & r & ":M" & r).Font.Bold = True  ' BOLD toan bo dong
                wsTB.Range("A" & r).Interior.Pattern = xlNone
                wsTB.Range("K" & r & ":M" & r).Interior.Pattern = xlNone
            Else
                wsTB.Range("A" & r & ":M" & r).Interior.Pattern = xlNone
            End If
        End If
    Next r
    ' Tong hang 1 (SUBTOTAL de ho tro filter)
    wsTB.Cells(1, 5).Formula = "=SUBTOTAL(9,E4:E" & lastRowTB & ")"
    wsTB.Cells(1, 6).Formula = "=SUBTOTAL(9,F4:F" & lastRowTB & ")"
    wsTB.Cells(1, 7).Formula = "=SUBTOTAL(9,G4:G" & lastRowTB & ")"
    wsTB.Cells(1, 8).Formula = "=SUBTOTAL(9,H4:H" & lastRowTB & ")"
    wsTB.Cells(1, 9).Formula = "=SUBTOTAL(9,I4:I" & lastRowTB & ")"
    wsTB.Cells(1, 10).Formula = "=SUBTOTAL(9,J4:J" & lastRowTB & ")"
    wsTB.Cells(1, 12).ClearContents
    wsTB.Cells(1, 13).ClearContents
    wsTB.Range("E1:L" & lastRowTB).NumberFormat = "#,##0"
    Auto_Tinh_TB = ""
    Exit Function
ErrorHandler:
    Application.Calculation = oldCalc
    MsgBox "L" & ChrW(7894) & "I: Kh" & ChrW(244) & "ng th" & ChrW(7875) & " t" & ChrW(237) & "nh to" & ChrW(225) & "n TB!" & vbCrLf & vbCrLf & _
           "Chi ti" & ChrW(7871) & "t: " & Err.Description, vbCritical
    Auto_Tinh_TB = ""
End Function

Private Function NormalizeTBNumberColumns(wsTB As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As String
    On Error Resume Next
    If wsTB Is Nothing Then Exit Function
    If lastRow < firstRow Then Exit Function

    Dim rng As Range
    Set rng = wsTB.Range("E" & firstRow & ":J" & lastRow)
    Dim arr As Variant
    arr = rng.Value

    Dim r As Long, c As Long
    Dim tmp As String
    Dim badCount As Long
    Dim badList As String

    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If IsEmpty(arr(r, c)) Then GoTo NextCell
            If IsNumeric(arr(r, c)) Then GoTo NextCell
            tmp = Trim$(CStr(arr(r, c)))
            If Len(tmp) = 0 Then
                arr(r, c) = ""
            ElseIf IsNumeric(tmp) Then
                arr(r, c) = Val(tmp)
            Else
                arr(r, c) = 0
                badCount = badCount + 1
                If badCount <= 5 Then
                    badList = badList & IIf(Len(badList) > 0, ", ", "") & rng.Cells(r, c).Address(False, False)
                End If
            End If
NextCell:
        Next c
    Next r

    rng.Value = arr
    rng.NumberFormat = "#,##0"

    If badCount > 0 Then
        NormalizeTBNumberColumns = "TB: " & badCount & " ô E:J không phải số (đã đặt =0). Kiểm tra các ô: " & badList
    End If
End Function
' Wrapper cho Ribbon button - Cap nhat dropdown TH
Public Sub Update_TH_Dropdown_Button(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsTH As Worksheet, wsTB As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsTH = GetSheet(wb, "TH")
    Set wsTB = GetSheet(wb, "TB")
    If wsTH Is Nothing Or wsTB Is Nothing Then
        MsgBox "Can co sheet TH va TB truoc khi cap nhat dropdown.", vbExclamation
        Exit Sub
    End If
    Update_TH_Dropdown
    MsgBox "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t dropdown TH th" & ChrW(224) & "nh c" & ChrW(244) & "ng!", vbInformation
End Sub
' Cap nhat dropdown cho C4 dua tren TK goc trong B4
Public Sub Update_TH_Dropdown()
    Dim wb As Workbook
    Dim wsTH As Worksheet, wsTB As Worksheet
    Dim tkPrefix As String
    Dim dictTK As Object
    Dim lastRowTB As Long, r As Long
    Dim tkList As String, tkItem As String
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsTH = wb.Sheets("TH")
    Set wsTB = wb.Sheets("TB")
    On Error GoTo 0
    If wsTH Is Nothing Or wsTB Is Nothing Then Exit Sub
    ' Lay TK prefix tu B4
    tkPrefix = Trim$(CStr(wsTH.Range("B4").Value))
    ' Loai bo apostrophe neu co
    If Left$(tkPrefix, 1) = "'" Then tkPrefix = Mid$(tkPrefix, 2)
    tkPrefix = Trim$(tkPrefix)
    If tkPrefix = "" Then
        ' Xoa validation neu B4 trong
        On Error Resume Next
        wsTH.Range("C4").Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If
    ' Thu thap tat ca TK con tu TB
    Set dictTK = CreateObject("Scripting.Dictionary")
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
    For r = 4 To lastRowTB
        tkItem = Trim$(CStr(wsTB.Cells(r, 3).Value))
        ' Loai bo apostrophe neu co
        If Left$(tkItem, 1) = "'" Then tkItem = Mid$(tkItem, 2)
        tkItem = Trim$(tkItem)
        If tkItem <> "" And Left$(tkItem, Len(tkPrefix)) = tkPrefix Then
            If Not dictTK.Exists(tkItem) Then
                dictTK.Add tkItem, True
            End If
        End If
    Next r
    ' Tao list cho validation
    If dictTK.Count > 0 Then
        Dim keys As Variant, i As Long
        keys = dictTK.keys
        tkList = keys(0)
        For i = 1 To UBound(keys)
            tkList = tkList & "," & keys(i)
        Next i
        ' Neu list qua dai (>255 ky tu), dung Named Range thay the
        If Len(tkList) > 255 Then
            ' Ghi list vao hidden area va dung Named Range
            Dim helperRange As Range
            Set helperRange = wsTH.Range("Z1").Resize(dictTK.Count, 1)
            helperRange.ClearContents
            For i = 0 To UBound(keys)
                helperRange.Cells(i + 1, 1).Value = keys(i)
            Next i
            ' Ap dung validation dung Named Range
            With wsTH.Range("C4").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                     Formula1:="=" & helperRange.Address
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Else
            ' Ap dung validation truc tiep
            With wsTH.Range("C4").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                     Operator:=xlBetween, Formula1:=tkList
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
    Else
        ' Khong co TK con -> xoa validation
        On Error Resume Next
        wsTH.Range("C4").Validation.Delete
        On Error GoTo 0
    End If
End Sub
Public Function Auto_Tinh_TH(wsNKC As Worksheet) As String
    Dim wb As Workbook
    Dim wsTH As Worksheet, wsTB As Worksheet, wsData As Worksheet
    Dim tkRaw As String, tkRoot As String
    Dim lenMain As Long, lenOpp As Long
    Dim lastNKC As Long, lastTB As Long, lastData As Long
    Dim r As Long
    Dim monthFilter As Long, hasMonthFilter As Boolean
    Dim slotCount As Long
    Dim duNoDK As Double, duCoDK As Double
    Dim tkNoFull As String, tkCoFull As String, soTien As Double
    Dim oppKey As String
    Dim dictOpp4 As Object
    Dim dictDebit As Object, dictCredit As Object
    Dim dictDebitFull As Object, dictCreditFull As Object
    Dim pairs As Variant
    Dim i As Long
    Dim oppLenSetting As Long
    Dim totalDebitPS As Double, totalCreditPS As Double
    Dim sdBalance As Double
    Dim filterState As Collection
    Dim warnNonNum As String
    On Error GoTo ErrHandler
    Set wb = wsNKC.Parent
    On Error Resume Next
    Set wsTH = wb.Sheets("TH")
    Set wsTB = wb.Sheets("TB")
    Set wsData = wb.Sheets("Data")
    On Error GoTo ErrHandler
    ' Neu chua co TH thi tao moi (sau khi NKC/TB da co du lieu)
    If wsTH Is Nothing Then
        On Error Resume Next
        If wb.Worksheets.Count >= 2 Then
            Set wsTH = Tao_TH_Template(wb, wb.Worksheets(wb.Worksheets.Count))
        Else
            Set wsTH = Tao_TH_Template(wb, wsNKC)
        End If
        On Error GoTo ErrHandler
        ' Đảm bảo event TH được gán sau khi tạo mới
        On Error Resume Next
        Application.Run "Enable_TH_AutoRefresh"
        On Error GoTo ErrHandler
    End If
    If wsTH Is Nothing Then
        Auto_Tinh_TH = "Khong the tao sheet TH."
        Exit Function
    End If
    ' Luu trang thai filter NKC va tam thoi bo loc de tinh tren full data
    Set filterState = CaptureAutoFilterState(wsNKC)
    On Error Resume Next
    If wsNKC.FilterMode Then wsNKC.ShowAllData
    On Error GoTo ErrHandler
    tkRaw = NormalizeAccount(wsTH.Range("C4").Value)
    If tkRaw = "" Then
        Auto_Tinh_TH = "" ' silent if chua nhap TK
        Exit Function
    End If
    ' So dong doi ung hien thi - se tinh dong theo so luong thuc te
    ' (tam thoi gan = 0, se cap nhat sau khi biet count thuc te)
    slotCount = 0
    ' D3: cap TK doi ung (so ky tu LEFT), khong phai so dong
    oppLenSetting = 0
    If IsNumeric(wsTH.Range("D3").Value) Then
        oppLenSetting = CLng(wsTH.Range("D3").Value)
        If oppLenSetting < 0 Then oppLenSetting = 0
    End If
    If IsNumeric(wsTH.Range("D2").Value) Then
        monthFilter = CLng(wsTH.Range("D2").Value)
        hasMonthFilter = (monthFilter > 0)
    End If
    Set dictOpp4 = CreateObject("Scripting.Dictionary")
    If Not wsData Is Nothing Then
        lastData = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
        For r = 1 To lastData
            If Len(NormalizeAccount(wsData.Cells(r, "L").Value)) >= 4 Then
                dictOpp4(Left$(NormalizeAccount(wsData.Cells(r, "L").Value), 4)) = True
            End If
        Next r
    End If
    ' Dinh nghia TK goc theo do dai nguoi dung nhap (khong ep tu Data)
    lenMain = Len(tkRaw)
    If lenMain < 1 Then
        Auto_Tinh_TH = "Sheet TH: vui long nhap tai khoan tai o C4."
        Exit Function
    End If
    tkRoot = Left$(tkRaw, lenMain)
    ' Clear vung du lieu cu de tranh sot dong khi doi cap TK
    ClearTHDataArea wsTH
    ' Tinh so du dau ky tu TB neu co (uu tien khop dung cap TK, tranh double TK cap 3/4)
    If Not wsTB Is Nothing Then
        Dim duNoExact As Double, duCoExact As Double
        Dim duNoLeft As Double, duCoLeft As Double
        Dim hasExact As Boolean
        lastTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
        For r = 4 To lastTB
            Dim tkTB As String
            tkTB = NormalizeAccount(wsTB.Cells(r, 3).Value)
            If tkTB <> "" And Left$(tkTB, lenMain) = tkRoot Then
                If Len(tkTB) = lenMain Then
                    duNoExact = duNoExact + CDbl(val(wsTB.Cells(r, 5).Value))
                    duCoExact = duCoExact + CDbl(val(wsTB.Cells(r, 6).Value))
                    hasExact = True
                Else
                    duNoLeft = duNoLeft + CDbl(val(wsTB.Cells(r, 5).Value))
                    duCoLeft = duCoLeft + CDbl(val(wsTB.Cells(r, 6).Value))
                End If
            End If
        Next r
        If hasExact Then
            duNoDK = duNoExact
            duCoDK = duCoExact
        Else
            duNoDK = duNoLeft
            duCoDK = duCoLeft
        End If
    End If
    ' Thu thap phat sinh tu NKC
    Set dictDebit = CreateObject("Scripting.Dictionary")
    Set dictCredit = CreateObject("Scripting.Dictionary")
    Set dictDebitFull = CreateObject("Scripting.Dictionary")
    Set dictCreditFull = CreateObject("Scripting.Dictionary")
    lastNKC = wsNKC.Cells(wsNKC.Rows.Count, "A").End(xlUp).Row
    For r = 3 To lastNKC
        If Not hasMonthFilter Or wsNKC.Cells(r, 3).Value = monthFilter Then
            tkNoFull = NormalizeAccount(wsNKC.Cells(r, 8).Value)
            tkCoFull = NormalizeAccount(wsNKC.Cells(r, 9).Value)
            soTien = CDbl(val(wsNKC.Cells(r, 10).Value))
            If tkNoFull <> "" And Left$(tkNoFull, lenMain) = tkRoot Then
                If oppLenSetting > 0 And oppLenSetting >= 4 Then
                    oppKey = tkCoFull ' lay full neu yeu cau >=4
                Else
                    lenOpp = GetPrefixLenFromDict(tkCoFull, dictOpp4, CLng(oppLenSetting))
                    oppKey = Left$(tkCoFull, lenOpp)
                End If
                DictAddSumWithFull dictDebit, dictDebitFull, oppKey, tkCoFull, soTien
                totalDebitPS = totalDebitPS + soTien
            ElseIf tkCoFull <> "" And Left$(tkCoFull, lenMain) = tkRoot Then
                If oppLenSetting > 0 And oppLenSetting >= 4 Then
                    oppKey = tkNoFull
                Else
                    lenOpp = GetPrefixLenFromDict(tkNoFull, dictOpp4, CLng(oppLenSetting))
                    oppKey = Left$(tkNoFull, lenOpp)
                End If
                DictAddSumWithFull dictCredit, dictCreditFull, oppKey, tkNoFull, soTien
                totalCreditPS = totalCreditPS + soTien
            End If
        End If
    Next r

    ' Ghi so du dau ky
    wsTH.Range("C5").Value = duNoDK
    wsTH.Range("D5").Value = duCoDK

    ' Check xem co bat Rut gon khong (D4)
    Dim enableGrouping As Boolean
    enableGrouping = True  ' Mac dinh
    On Error Resume Next
    If wsTH.Range("D4").Value = False Then enableGrouping = False
    On Error GoTo ErrHandler

    ' Xu ly va ghi phat sinh No/Co, sau do tinh actualSlots
    Dim maxDebit As Long, maxCredit As Long
    maxDebit = 0
    maxCredit = 0

    ' Ghi phat sinh No (TK goc o ben No -> doi ung o col B/C)
    If dictDebit.Count > 0 Then
        pairs = SortDictByAbsWithFull(dictDebit, dictDebitFull)
        ' Kiem tra xem co TK dai (>=6) khong, neu co thi nhom theo cap 3 (neu bat Rut gon)
        Dim hasLongDebit As Boolean
        hasLongDebit = CheckHasLongTK(dictDebitFull)

        If hasLongDebit And enableGrouping Then
            ' Nhom lai theo cap 3 (GroupByLevel3 da sort san)
            pairs = GroupByLevel3(pairs)
        End If

        maxDebit = UBound(pairs, 1)
        For i = 1 To maxDebit
            Dim dispOppN As String
            dispOppN = pairs(i, 1)
            wsTH.Cells(5 + i, 2).Value = IIf(dispOppN <> "", "<" & dispOppN & ">", "")
            wsTH.Cells(5 + i, 3).Value = pairs(i, 2) ' so tien
        Next i
    End If

    ' Ghi phat sinh Co (TK goc o ben Co -> doi ung o col D/E)
    If dictCredit.Count > 0 Then
        pairs = SortDictByAbsWithFull(dictCredit, dictCreditFull)
        ' Kiem tra xem co TK dai (>=6) khong, neu co thi nhom theo cap 3 (neu bat Rut gon)
        Dim hasLongCredit As Boolean
        hasLongCredit = CheckHasLongTK(dictCreditFull)

        If hasLongCredit And enableGrouping Then
            ' Nhom lai theo cap 3 (GroupByLevel3 da sort san)
            pairs = GroupByLevel3(pairs)
        End If

        maxCredit = UBound(pairs, 1)
        For i = 1 To maxCredit
            Dim dispOppC As String
            dispOppC = pairs(i, 1)
            wsTH.Cells(5 + i, 4).Value = pairs(i, 2) ' so tien
            wsTH.Cells(5 + i, 5).Value = IIf(dispOppC <> "", "<" & dispOppC & ">", "")
        Next i
    End If

    ' Tinh actualSlots SAU KHI da group (dua vao so dong thuc te)
    Dim actualSlots As Long
    actualSlots = Application.Max(maxDebit, maxCredit)
    If actualSlots < 1 Then actualSlots = 1

    ' Tinh vi tri dong SPS va SDCK
    Dim rowSPS As Long, rowSDCK As Long
    rowSPS = 5 + actualSlots + 1
    rowSDCK = rowSPS + 1

    ' Clear vung du lieu cu phia duoi (dong thua)
    Dim lastClearRow As Long
    Dim lastExisting As Long
    lastExisting = wsTH.Cells(wsTH.Rows.Count, "A").End(xlUp).Row
    lastExisting = Application.Max(lastExisting, wsTH.Cells(wsTH.Rows.Count, "E").End(xlUp).Row)
    lastClearRow = Application.Max(rowSDCK + 20, lastExisting)

    ' Clear dong thua giua data va SPS
    ' Data rows: 6 to (5 + actualSlots), SPS row: (5 + actualSlots + 1) = (6 + actualSlots)
    ' No gap exists, so no clearing needed here
    ' (Previous code incorrectly cleared the last data row)

    ' Clear dong thua phia duoi SDCK
    wsTH.Range("A" & (rowSDCK + 1) & ":E" & lastClearRow).ClearContents
    wsTH.Range("B" & (rowSDCK + 1) & ":E" & lastClearRow).Interior.Pattern = xlNone
    wsTH.Range("B" & (rowSDCK + 1) & ":E" & lastClearRow).Borders.LineStyle = xlNone
    wsTH.Range("I7:K" & (lastClearRow + 1)).ClearContents

    ' Format khu vuc phat sinh (background + border)
    Dim firstDataRow As Long, lastDataRow As Long
    firstDataRow = 6
    lastDataRow = 5 + actualSlots
    wsTH.Range("B" & firstDataRow & ":E" & lastDataRow).Interior.Color = RGB(242, 242, 242)
    wsTH.Range("B" & firstDataRow & ":E" & lastDataRow).Borders.LineStyle = xlContinuous
    wsTH.Range("B" & firstDataRow & ":E" & lastDataRow).Borders.Color = RGB(200, 200, 200)

    ' Ghi nhan SPS va SDCK o vi tri dong
    wsTH.Cells(rowSPS, 1).Value = "SPS"
    wsTH.Cells(rowSPS, 1).Font.Bold = True
    wsTH.Cells(rowSPS, 1).Font.Color = RGB(0, 0, 200)
    wsTH.Cells(rowSDCK, 1).Value = "SDCK"
    wsTH.Cells(rowSDCK, 1).Font.Bold = True
    wsTH.Cells(rowSDCK, 1).Font.Color = RGB(0, 0, 200)

    ' Ghi SPS va SDCK theo T-account
    wsTH.Cells(rowSPS, 3).Value = totalDebitPS
    wsTH.Cells(rowSPS, 4).Value = totalCreditPS
    sdBalance = (duNoDK - duCoDK) + (totalDebitPS - totalCreditPS)
    wsTH.Cells(rowSDCK, 3).Value = Application.Max(sdBalance, 0)
    wsTH.Cells(rowSDCK, 4).Value = Application.Max(-sdBalance, 0)
    ' Khoi phuc filter NKC (neu co)
    RestoreAutoFilterState wsNKC, filterState
    Auto_Tinh_TH = ""
    Exit Function
ErrHandler:
    RestoreAutoFilterState wsNKC, filterState
    Auto_Tinh_TH = "Khong the cap nhat sheet TH. Chi tiet: " & Err.Description
End Function
Private Function CaptureAutoFilterState(ws As Worksheet) As Collection
    Dim af As AutoFilter
    Dim state As Collection
    Dim i As Long
    On Error Resume Next
    Set af = ws.AutoFilter
    If af Is Nothing Then Exit Function
    Set state = New Collection
    For i = 1 To af.Filters.Count
        If af.Filters(i).On Then
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            dict("Field") = i
            dict("Operator") = af.Filters(i).Operator
            On Error Resume Next
            dict("Criteria1") = af.Filters(i).Criteria1
            dict("Criteria2") = af.Filters(i).Criteria2
            On Error GoTo 0
            state.Add dict
        End If
    Next i
    Set CaptureAutoFilterState = state
End Function
Private Sub RestoreAutoFilterState(ws As Worksheet, state As Collection)
    Dim afRange As Range
    Dim item As Variant
    If state Is Nothing Then Exit Sub
    On Error Resume Next
    Set afRange = ws.AutoFilter.Range
    If afRange Is Nothing Then Exit Sub
    For Each item In state
        If item.Exists("Criteria2") And Not IsEmpty(item("Criteria2")) Then
            afRange.AutoFilter Field:=item("Field"), Criteria1:=item("Criteria1"), _
                               Operator:=item("Operator"), Criteria2:=item("Criteria2")
        Else
            afRange.AutoFilter Field:=item("Field"), Criteria1:=item("Criteria1")
        End If
    Next item
End Sub
Private Function NormalizeAccount(v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then
        NormalizeAccount = ""
        Exit Function
    End If
    If Right$(s, 2) = ".0" Then s = Left$(s, Len(s) - 2)
    NormalizeAccount = s
End Function
Private Sub ClearTHDataArea(wsTH As Worksheet)
    Dim lastB As Long, lastE As Long, lastA As Long, lastRow As Long
    lastB = wsTH.Cells(wsTH.Rows.Count, "B").End(xlUp).Row
    lastE = wsTH.Cells(wsTH.Rows.Count, "E").End(xlUp).Row
    lastA = wsTH.Cells(wsTH.Rows.Count, "A").End(xlUp).Row
    lastRow = Application.Max(lastB, lastE, lastA, 200)
    If lastRow < 6 Then lastRow = 6
    wsTH.Range("A6:E" & lastRow).ClearContents
    wsTH.Range("B6:E" & lastRow).Interior.Pattern = xlNone
    wsTH.Range("B6:E" & lastRow).Borders.LineStyle = xlNone
    wsTH.Range("I7:K" & (lastRow + 1)).ClearContents
End Sub
Private Function GetPrefixLenFromDict(tk As String, dict4 As Object, Optional overrideLen As Long = 0) As Long
    Dim safeLen As Long
    safeLen = Len(tk)
    If safeLen = 0 Then
        GetPrefixLenFromDict = 0
        Exit Function
    End If
    If overrideLen > 0 Then
        If overrideLen > safeLen Then overrideLen = safeLen
        GetPrefixLenFromDict = overrideLen
        Exit Function
    End If
    If safeLen >= 4 And Not dict4 Is Nothing Then
        If dict4.Exists(Left$(tk, 4)) Then
            GetPrefixLenFromDict = 4
            Exit Function
        End If
    End If
    If safeLen < 3 Then
        GetPrefixLenFromDict = safeLen
    Else
        GetPrefixLenFromDict = 3
    End If
End Function
Private Sub DictAddSumWithFull(dictSum As Object, dictFull As Object, key As String, fullKey As String, val As Double)
    If key = "" Then Exit Sub
    If dictSum.Exists(key) Then
        dictSum(key) = dictSum(key) + val
        ' Luon luu full TK dai nhat de phat hien TK cap 3/4
        If dictFull.Exists(key) Then
            If Len(fullKey) > Len(CStr(dictFull(key))) Then
                dictFull(key) = fullKey
            End If
        Else
            dictFull.Add key, fullKey
        End If
    Else
        dictSum.Add key, val
        dictFull.Add key, fullKey
    End If
End Sub
Private Function SortDictByAbsWithFull(dictSum As Object, dictFull As Object) As Variant
    Dim n As Long, keys As Variant
    Dim vals() As Double, fullKeys() As String
    Dim i As Long, j As Long
    Dim tmpVal As Double, tmpKey As Variant, tmpFull As String
    n = dictSum.Count
    If n = 0 Then
        SortDictByAbsWithFull = Array()
        Exit Function
    End If
    keys = dictSum.keys
    ReDim vals(0 To n - 1)
    ReDim fullKeys(0 To n - 1)
    For i = 0 To n - 1
        vals(i) = dictSum(keys(i))
        If dictFull.Exists(keys(i)) Then
            fullKeys(i) = dictFull(keys(i))
        Else
            fullKeys(i) = keys(i)
        End If
    Next i
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If Abs(vals(j)) > Abs(vals(i)) Then
                tmpVal = vals(i)
                vals(i) = vals(j)
                vals(j) = tmpVal
                tmpKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tmpKey
                tmpFull = fullKeys(i)
                fullKeys(i) = fullKeys(j)
                fullKeys(j) = tmpFull
            End If
        Next j
    Next i
    Dim res() As Variant
    ReDim res(1 To n, 1 To 3)
    For i = 0 To n - 1
        res(i + 1, 1) = keys(i)
        res(i + 1, 2) = vals(i)
        res(i + 1, 3) = fullKeys(i)
    Next i
    SortDictByAbsWithFull = res
End Function
' ==================================================================================
' KIEM TRA CAP TAI KHOAN HOP LE THEO QUY TAC KE TOAN VIET NAM
' ==================================================================================
Private Function IsValidAccountPairCached(tkNo As String, tkCo As String, cache As Object) As Boolean
    Dim key As String
    If cache Is Nothing Then
        IsValidAccountPairCached = IsValidAccountPair(tkNo, tkCo)
        Exit Function
    End If
    key = tkNo & "|" & tkCo
    If cache.Exists(key) Then
        IsValidAccountPairCached = cache(key)
    Else
        IsValidAccountPairCached = IsValidAccountPair(tkNo, tkCo)
        cache.Add key, IsValidAccountPairCached
    End If
End Function

Function IsValidAccountPair(tkNo As String, tkCo As String) As Boolean
    Dim noPrefix As String, coPrefix As String
    noPrefix = Left(tkNo, 3)
    coPrefix = Left(tkCo, 3)
    IsValidAccountPair = False

    ' ==================================================================================
    ' 1. KET CHUYEN 911 (Thong tu 200)
    ' ==================================================================================
    ' Ket chuyen doanh thu: 5xx, 7xx No / 911 Co
    If (Left(tkNo, 1) = "5" Or Left(tkNo, 1) = "7") And coPrefix = "911" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ket chuyen chi phi: 911 No / 6xx, 8xx Co
    If noPrefix = "911" And (Left(tkCo, 1) = "6" Or Left(tkCo, 1) = "8") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Xac dinh ket qua: 911 <-> 421
    If (noPrefix = "911" And coPrefix = "421") Or (noPrefix = "421" And coPrefix = "911") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 2. QUY TAC MUA HANG (Nguy�n v?t li?u, h�ng h�a, TSC�)
    ' ==================================================================================
    ' Mua NVL, CCDC, h�ng h�a: 152, 153, 156 N? / 111, 112, 331 C�
    'N? 151 c� 111,112,331
    If (noPrefix = "152" Or noPrefix = "153" Or noPrefix = "156" Or noPrefix = "151") And _
       (coPrefix = "331" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If


    ' Mua TSC�: 211, 213 N? / 111, 112, 331 C�
    If (noPrefix = "211" Or noPrefix = "213") And _
       (coPrefix = "331" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Nh?n g�p v?n TSC�: 211 N? / 411 C�
    If noPrefix = "211" And coPrefix = "411" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Mua B�S�T: 217 N? / 111, 112, 331 C�
    If noPrefix = "217" And (coPrefix = "331" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 3. QUY TAC BAN HANG (Doanh thu va cung cap dich vu)
    ' ==================================================================================
    ' Doanh thu: 131, 111, 112 No / 511, 711 Co
    If (noPrefix = "131" Or noPrefix = "111" Or noPrefix = "112") And _
       (coPrefix = "511" Or coPrefix = "711") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Gia von hang ban: 632 No / 154, 155, 156 Co
    If noPrefix = "632" And (coPrefix = "154" Or coPrefix = "155" Or coPrefix = "156") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Gia von hang hoa: 632 No / 156 Co
    If noPrefix = "632" And coPrefix = "156" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Giam tru doanh thu: 521 No / 111, 112, 131 Co
    If noPrefix = "521" And (coPrefix = "131" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ket chuyen giam tru doanh thu: 511 No / 521 Co
    If noPrefix = "511" And coPrefix = "521" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Nhap lai hang tra lai: 156 No / 632 Co
    If noPrefix = "156" And coPrefix = "632" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 4. QUY TAC THUE GTGT (Thu? gi� tr? gia tang)
    ' ==================================================================================
    ' Thu? GTGT d?u v�o: 133 N? / 111, 112, 331 C�
    If noPrefix = "133" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "331") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu? GTGT du?c kh?u tr? (Th�ng tu 99): 133 N? / 331, 111, 112 C�
    If noPrefix = "133" And coPrefix = "331" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu? GTGT d?u ra: 131, 111, 112 N? / 333 C�
    If (noPrefix = "131" Or noPrefix = "111" Or noPrefix = "112") And coPrefix = "333" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu? GTGT ph?i n?p (kh�ng du?c kh?u tr?): 333 N? / 111, 112, 331 C�
    If noPrefix = "333" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "331") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 5. QUY TAC THANH TOAN (Ti?n m?t, ti?n g?i, c�ng n?)
    ' ==================================================================================
    ' Tr? ti?n ngu?i b�n: 331 N? / 111, 112 C�
    If noPrefix = "331" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu ti?n kh�ch h�ng: 111, 112 N? / 131 C�
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "131" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Chuy?n d?i ti?n: 111 <-> 112
    If (noPrefix = "111" And coPrefix = "112") Or (noPrefix = "112" And coPrefix = "111") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 6. QUY TAC LUONG & BAO HIEM (Luong, BHXH, BHYT)
    ' ==================================================================================
    ' Tr�ch luong ph?i tr?: 622, 627, 641, 642 N? / 334 C�
    If (noPrefix = "622" Or noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And _
       coPrefix = "334" Or coPrefix = "338" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Tr? luong: 334 N? / 111, 112 C�
    If noPrefix = "334" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Tr�ch BHXH, BHYT: 334 N? / 338 C�
    If noPrefix = "334" And coPrefix = "338" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' N?p BHXH: 338 N? / 111, 112 C�
    If noPrefix = "338" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 7. QUY TAC KHAU HAO (Kh?u hao TSC�)
    ' ==================================================================================
    ' Tr�ch kh?u hao: 627, 641, 642 N? / 214 C�
    If (noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And coPrefix = "214" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 8. QUY TAC VAY (Vay va nghia vu phai tra)
    ' ==================================================================================
    ' Nhan vay/thu tu cho vay tai chinh: 111, 112 No / 341 Co
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "341" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Tra vay: 341 No / 111, 112 Co
    If noPrefix = "341" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Phat hanh trai phieu: 111, 112 No / 343 Co
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "343" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thanh toan trai phieu: 343 No / 111, 112 Co
    If noPrefix = "343" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Nhan truoc dai han cua khach hang: 111, 112 No / 344 Co
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "344" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Hoan tra nhan truoc dai han: 344 No / 111, 112 Co
    If noPrefix = "344" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Trich chi phi phai tra: 627, 635, 641, 642 No / 335 Co
    If (noPrefix = "627" Or noPrefix = "635" Or noPrefix = "641" Or noPrefix = "642") And coPrefix = "335" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Dac thu: 333 No / 335 Co
    If noPrefix = "335" And coPrefix = "333" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thanh toan chi phi phai tra: 335 No / 111, 112, 331 Co
    If noPrefix = "335" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "331") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 9. QUY TAC DAUTU (�?u tu t�i ch�nh)
    ' ==================================================================================
    ' �?u tu ng?n h?n: 121, 128 N? / 111, 112 C�
    If (noPrefix = "121" Or noPrefix = "128") And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu h?i d?u tu ng?n h?n: 111, 112 N? / 121, 128 C�
    If (noPrefix = "111" Or noPrefix = "112") And (coPrefix = "121" Or coPrefix = "128") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' �?u tu d�i h?n: 221, 222, 228 N? / 111, 112, 411 C�
    If (noPrefix = "221" Or noPrefix = "222" Or noPrefix = "228") And _
       (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "411") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu h?i d?u tu d�i h?n: 111, 112 N? / 221, 222, 228 C�
    If (noPrefix = "111" Or noPrefix = "112") And _
       (coPrefix = "221" Or coPrefix = "222" Or coPrefix = "228") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 10. QUY TAC UNG TRUOC (T?m ?ng, ?ng tru?c)
    ' ==================================================================================
    ' T?m ?ng: 141 N? / 111, 112 C�
    If noPrefix = "141" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ho�n ?ng, thanh to�n t?m ?ng: 111, 112, 622, 627, 641, 642 N? / 141 C�
    If (noPrefix = "111" Or noPrefix = "112" Or noPrefix = "622" Or noPrefix = "627" Or _
        noPrefix = "641" Or noPrefix = "642") And coPrefix = "141" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Nh?n ?ng tru?c: 111, 112 N? / 131 C� (ghi tang c�ng n? ph?i thu d?ng th?i)
    ' (�� c� trong quy t?c thanh to�n)

    ' ==================================================================================
    ' 11. QUY TAC CHI PHI TRA TRUOC (Tr? tru?c ng?n h?n, d�i h?n)
    ' ==================================================================================
    ' Chi ph� tr? tru?c ng?n h?n: 142 N? / 111, 112, 331 C�
    If noPrefix = "142" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "331") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ph�n b? chi ph� tr? tru?c ng?n h?n: 622, 627, 641, 642 N? / 142 C�
    If (noPrefix = "622" Or noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And coPrefix = "142" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Chi ph� tr? tru?c d�i h?n: 242, 244 N? / 111, 112, 331 C�
    If (noPrefix = "242" Or noPrefix = "244") And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "331") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ph�n b? chi ph� tr? tru?c d�i h?n: 627, 641, 642 N? / 242, 244 C�
    If (noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And (coPrefix = "242" Or coPrefix = "244") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 12. QUY TAC VON CHU SO HUU (V?n, l?i nhu?n chua ph�n ph?i)
    ' ==================================================================================
    ' G�p v?n: 111, 112, 152, 156, 211 N? / 411 C�
    If (noPrefix = "111" Or noPrefix = "112" Or noPrefix = "152" Or noPrefix = "156" Or noPrefix = "211") And _
       coPrefix = "411" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' R�t v?n: 411 N? / 111, 112 C�
    If noPrefix = "411" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Tang v?n t? l?i nhu?n: 421 N? / 411 C�
    If noPrefix = "421" And coPrefix = "411" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Chia l?i nhu?n: 421 N? / 111, 112, 334 C�
    If noPrefix = "421" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "334") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Tr�ch qu?: 421 N? / 414, 418 C�
    If noPrefix = "421" And (coPrefix = "414" Or coPrefix = "418") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' S? d?ng qu?: 414, 418 N? / 111, 112, 211 C�
    If (noPrefix = "414" Or noPrefix = "418") And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "211") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 13. QUY TAC SAN XUAT (Chi ph� s?n xu?t, gi� th�nh)
    ' ==================================================================================
    ' Xu?t NVL s?n xu?t: 621, 154 N? / 152 C�
    If (noPrefix = "621" Or noPrefix = "154") And coPrefix = "152" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Xu?t CCDC s?n xu?t: 622, 627 N? / 153 C�
    If (noPrefix = "622" Or noPrefix = "627") And coPrefix = "153" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' K?t chuy?n chi ph� s?n xu?t: 154 N? / 621, 622, 627 C�
    If noPrefix = "154" And (coPrefix = "621" Or coPrefix = "622" Or coPrefix = "627") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Nh?p th�nh ph?m: 155 N? / 154 C�
    If noPrefix = "155" And coPrefix = "154" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 14. QUY TAC PHAI THU/TRA KHAC (Ph?i thu kh�c, ph?i tr? kh�c)
    ' ==================================================================================
    ' Ph?i thu kh�c: 138 N? / 111, 112, 711 C�
    If noPrefix = "138" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "711") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Thu ph?i thu kh�c: 111, 112 N? / 138 C�
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "138" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ph?i tr? kh�c: 338, 344 N? / 111, 112 C�
    If (noPrefix = "338" Or noPrefix = "344") And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ph?i thu v? b�n t�i s?n: 138 N? / 711 C�
    If noPrefix = "138" And coPrefix = "711" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' ==================================================================================
    ' 15. QUY TAC THONG TU 99/2024 (T�i kho?n m?i)
    ' ==================================================================================
    ' TK 171: Giao d?ch mua b�n l?i tr�i phi?u Ch�nh ph?
    ' If noPrefix = "171" And (coPrefix = "111" Or coPrefix = "112") Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "171" Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' ' TK 2281: Chi ph� ch? ph�n b? (CCDC ch? ph�n b?)
    ' If noPrefix = "2281" And (coPrefix = "331" Or coPrefix = "111" Or coPrefix = "112") Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' ' Ph�n b? CCDC: 627, 641, 642 N? / 2281 C�
    ' If (noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And coPrefix = "2281" Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' ' TK 229: D? ph�ng gi?m gi� h�ng t?n kho
    ' If (noPrefix = "632" Or noPrefix = "641") And (coPrefix = "229" Or Left(coPrefix, 3) = "229") Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' ' Ho�n nh?p d? ph�ng: 229 N? / 632, 711 C�
    ' If (coPrefix = "229" Or Left(coPrefix, 3) = "229") And (noPrefix = "632" Or noPrefix = "711") Then
    '     IsValidAccountPair = True: Exit Function
    ' End If

    ' ==================================================================================
    ' 16. CUNG TAI KHOAN (B�t to�n n?i b?)
    ' ==================================================================================
    If tkNo = tkCo Then
        IsValidAccountPair = True: Exit Function
    End If
End Function
' ===================================================================
' Ham Tinh Toan TB (Trial Balance)
' ===================================================================
Public Sub Tinh_Toan_TB(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wsNKC As Worksheet, wsTB As Worksheet, wsSource As Worksheet
    Dim wb As Workbook
    Dim lastRowNKC As Long, lastRowTB As Long, lastRowSource As Long
    Dim r As Long, i As Long
    Dim tkFull As String, tkTrim As String
    Dim dictSourceTK As Object
    Dim tkSource As String
    Dim useLeftMatch As Boolean
    Dim oldCalc As XlCalculation
    Set wb = ActiveWorkbook
    oldCalc = Application.Calculation
    ' Kiem tra sheet NKC ton tai
    If Not WorksheetExists("NKC", wb) Then
        MsgBox "Sheet NKC ch" & ChrW(432) & "a t" & ChrW(7891) & "n t" & ChrW(7841) & "i! H" & ChrW(227) & "y ch" & ChrW(7841) & "y Xu_Ly_NKC tr" & ChrW(432) & ChrW(7899) & "c.", vbExclamation
        Exit Sub
    End If
    ' Kiem tra sheet TB ton tai
    If Not WorksheetExists("TB", wb) Then
        MsgBox "Sheet TB ch" & ChrW(432) & "a t" & ChrW(7891) & "n t" & ChrW(7841) & "i! H" & ChrW(227) & "y ch" & ChrW(7841) & "y Tao_Mau_TB tr" & ChrW(432) & ChrW(7899) & "c.", vbExclamation
        Exit Sub
    End If
    Set wsNKC = wb.Sheets("NKC")
    Set wsTB = wb.Sheets("TB")
    ' Doc du lieu NKC
    lastRowNKC = wsNKC.Cells(wsNKC.Rows.Count, "A").End(xlUp).Row
    If lastRowNKC < 3 Then
        MsgBox "Sheet NKC kh" & ChrW(244) & "ng c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u!", vbExclamation
        Exit Sub
    End If
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
    If lastRowTB < 4 Then
        MsgBox "Sheet TB ch" & ChrW(432) & "a c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u!", vbExclamation
        Exit Sub
    End If
    ' Tao dictionary chua danh sach TK tu sheet "So Nhat Ky Chung"
    Set dictSourceTK = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsSource = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0
    If Not wsSource Is Nothing Then
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
        For i = 2 To lastRowSource
            tkSource = Trim$(CStr(wsSource.Cells(i, 4).Value))
            If tkSource <> "" And Not dictSourceTK.Exists(tkSource) Then
                dictSourceTK.Add tkSource, True
            End If
        Next i
    End If
    Application.Calculation = xlCalculationManual
    ' FORMAT COT C (TAI KHOAN) THANH TEXT DE TRANH BI CONVERT SANG NUMBER KHI EDIT
    wsTB.Columns(3).NumberFormat = "@"
    For r = 4 To lastRowTB
        If wsTB.Cells(r, 3).Value <> "" Then
            tkFull = CStr(wsTB.Cells(r, 3).Value)
            tkTrim = Trim$(tkFull)
            ' FORCE GHI LAI GIA TRI DE CHUYEN HOAN TOAN SANG TEXT (tranh bi luu duoi dang number)
            wsTB.Cells(r, 3).Value = "'" & tkTrim
            ' Luu do dai TK vao cot B (VALUE, khong phai formula de tranh loi khi user edit)
            wsTB.Cells(r, 2).Value = Len(tkTrim)
            ' Xac dinh logic: neu TK >= 6 ky tu va KHONG ton tai trong So Nhat Ky Chung => dung LEFT MATCH
            useLeftMatch = (Len(tkTrim) >= 6 And Not dictSourceTK.Exists(tkTrim))
            ' Cong thuc cho cot L (Lech No) - dung RC2 thay vi LEN(RC3) de tranh bi tinh lai khi edit
            If Len(tkTrim) = 3 Then
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C10:R" & lastRowNKC & "C10,NKC!R3C6:R" & lastRowNKC & "C6,RC3)-RC7"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C8:R" & lastRowNKC & "C8,RC2)=RC3)*NKC!R3C10:R" & lastRowNKC & "C10)-RC7"
            ElseIf useLeftMatch Then
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C8:R" & lastRowNKC & "C8,RC2)=RC3)*NKC!R3C10:R" & lastRowNKC & "C10)-RC7"
            Else
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C10:R" & lastRowNKC & "C10,NKC!R3C8:R" & lastRowNKC & "C8,RC3)-RC7"
            End If
            ' Cong thuc cho cot M (Lech Co) - dung RC2 thay vi LEN(RC3) de tranh bi tinh lai khi edit
            If Len(tkTrim) = 3 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C10:R" & lastRowNKC & "C10,NKC!R3C7:R" & lastRowNKC & "C7,RC3)-RC8"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C9:R" & lastRowNKC & "C9,RC2)=RC3)*NKC!R3C10:R" & lastRowNKC & "C10)-RC8"
            ElseIf useLeftMatch Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C9:R" & lastRowNKC & "C9,RC2)=RC3)*NKC!R3C10:R" & lastRowNKC & "C10)-RC8"
            Else
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C10:R" & lastRowNKC & "C10,NKC!R3C9:R" & lastRowNKC & "C9,RC3)-RC8"
            End If
        End If
    Next r
    wsTB.Calculate
    Application.Calculation = oldCalc
    ' To mau dong theo tieu chi
    For r = 4 To lastRowTB
        If wsTB.Cells(r, 3).Value <> "" Then
            Dim tkLen As Long
            tkLen = wsTB.Cells(r, 2).Value
            ' Uu tien 1: To vang dong co lech L hoac M
            If wsTB.Cells(r, 12).Value <> 0 Or wsTB.Cells(r, 13).Value <> 0 Then
                wsTB.Range("A" & r & ":M" & r).Interior.Color = RGB(255, 255, 150)
            ' Uu tien 2: To xanh (B den J) va BOLD cho dong co cap TK = 3
            ElseIf tkLen = 3 Then
                wsTB.Range("B" & r & ":J" & r).Interior.Color = RGB(146, 208, 80)  ' Xanh la nhat
                wsTB.Range("A" & r & ":M" & r).Font.Bold = True  ' BOLD toan bo dong
                wsTB.Range("A" & r).Interior.Pattern = xlNone
                wsTB.Range("K" & r & ":M" & r).Interior.Pattern = xlNone
            Else
                wsTB.Range("A" & r & ":M" & r).Interior.Pattern = xlNone
            End If
        End If
    Next r
    wsTB.Cells(1, 5).Formula = "=SUBTOTAL(9,E4:E" & lastRowTB & ")"  ' Tong Dau ky No
    wsTB.Cells(1, 6).Formula = "=SUBTOTAL(9,F4:F" & lastRowTB & ")"  ' Tong Dau ky Co
    wsTB.Cells(1, 7).Formula = "=SUBTOTAL(9,G4:G" & lastRowTB & ")"  ' Tong PS No
    wsTB.Cells(1, 8).Formula = "=SUBTOTAL(9,H4:H" & lastRowTB & ")"  ' Tong PS Co
    wsTB.Cells(1, 9).Formula = "=SUBTOTAL(9,I4:I" & lastRowTB & ")"  ' Tong Cuoi ky No
    wsTB.Cells(1, 10).Formula = "=SUBTOTAL(9,J4:J" & lastRowTB & ")"  ' Tong Cuoi ky Co
    ' Format number columns
    wsTB.Range("E1:L" & lastRowTB).NumberFormat = "#,##0"
    MsgBox "T" & ChrW(237) & "nh to" & ChrW(225) & "n TB th" & ChrW(224) & "nh c" & ChrW(244) & "ng! " & ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t c" & ChrW(244) & "ng th" & ChrW(7913) & "c B, L, M cho " & (lastRowTB - 3) & " d" & ChrW(242) & "ng.", vbInformation
End Sub
' Helper function to check if worksheet exists
Private Function WorksheetExists(ByVal sheetName As String, Optional wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = ActiveWorkbook
        On Error GoTo 0
        If wb Is Nothing Then Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
Private Function GetMonthValue(vDate As Variant) As Variant
    If IsDate(vDate) Then
        GetMonthValue = Month(vDate)
    Else
        GetMonthValue = ""
    End If
End Function

Private Function CheckHasLongTK(dictFull As Object) As Boolean
    ' Kiem tra xem trong dictFull co TK nao dai >= 6 ky tu khong
    Dim vals As Variant, keys As Variant, i As Long, tkVal As String
    CheckHasLongTK = False
    If dictFull Is Nothing Then Exit Function
    If dictFull.Count = 0 Then Exit Function

    ' Prefer checking full TK (dictFull stores fullKey as value)
    vals = dictFull.Items
    For i = 0 To UBound(vals)
        tkVal = CStr(vals(i))
        If Len(tkVal) >= 6 Then
            CheckHasLongTK = True
            Exit Function
        End If
    Next i

    ' Fallback: also check key length (in case Items is empty)
    keys = dictFull.keys
    For i = 0 To UBound(keys)
        tkVal = CStr(keys(i))
        If Len(tkVal) >= 6 Then
            CheckHasLongTK = True
            Exit Function
        End If
    Next i
End Function

Private Function GroupByLevel3(pairs As Variant) As Variant
    ' Nhom lai pairs theo cap 3 (rut gon TK ve 3 ky tu) va cong don so tien
    ' Input: pairs(i, 1) = key, pairs(i, 2) = amount, pairs(i, 3) = fullTK
    ' Output: pairs array voi (tk3, total_amount)
    Dim grouped As Object
    Dim i As Long, tkFull As String, tk3 As String, amount As Double
    Set grouped = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(pairs, 1)
        tkFull = CStr(pairs(i, 3)) ' full TK
        amount = CDbl(pairs(i, 2))

        ' Rut gon ve cap 3
        If Len(tkFull) >= 3 Then
            tk3 = Left$(tkFull, 3)
        Else
            tk3 = tkFull
        End If

        ' Cong don so tien theo nhom cap 3
        If grouped.Exists(tk3) Then
            grouped(tk3) = grouped(tk3) + amount
        Else
            grouped.Add tk3, amount
        End If
    Next i

    ' Convert dict thanh pairs array va sort
    GroupByLevel3 = DictToPairsSorted(grouped)
End Function

Private Function DictToPairsSorted(dict As Object) As Variant
    ' Convert Dictionary thanh pairs array va sort giam dan theo abs value
    ' Output: pairs(i, 1) = key, pairs(i, 2) = value
    Dim n As Long, keys As Variant, vals() As Double
    Dim i As Long, j As Long
    Dim tmpVal As Double, tmpKey As Variant
    Dim res() As Variant

    n = dict.Count
    If n = 0 Then
        ReDim res(1 To 1, 1 To 2)
        res(1, 1) = ""
        res(1, 2) = 0
        DictToPairsSorted = res
        Exit Function
    End If

    keys = dict.keys
    ReDim vals(0 To n - 1)

    ' Copy values
    For i = 0 To n - 1
        vals(i) = dict(keys(i))
    Next i

    ' Bubble sort giam dan theo abs
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If Abs(vals(j)) > Abs(vals(i)) Then
                tmpVal = vals(i)
                vals(i) = vals(j)
                vals(j) = tmpVal
                tmpKey = keys(i)
                keys(i) = keys(j)
                keys(j) = tmpKey
            End If
        Next j
    Next i

    ' Build result array
    ReDim res(1 To n, 1 To 2)
    For i = 0 To n - 1
        res(i + 1, 1) = keys(i)
        res(i + 1, 2) = vals(i)
    Next i

    DictToPairsSorted = res
End Function
