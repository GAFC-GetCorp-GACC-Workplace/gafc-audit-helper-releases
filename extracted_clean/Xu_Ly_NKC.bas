Attribute VB_Name = "Xu_Ly_NKC"
Option Explicit
Public Sub Test_Xu_ly_NKC()
    Xu_ly_NKC1111 Nothing
End Sub

' Bo sung cac cot thieu cho sheet NKC da paste
Private Sub Bo_Sung_Cot_NKC(wsNKC As Worksheet)
    Dim lastRow As Long, r As Long
    Dim tkNo As String, tkCo As String
    Dim amtVal As Variant
    Dim arr As Variant
    Dim ngayHT As Variant

    On Error Resume Next
    lastRow = GetLastUsedRow(wsNKC)
    On Error GoTo 0

    If lastRow < 3 Then Exit Sub

    ' Dam bao header du cot (them "Khac" va "Can review" neu thieu)
    EnsureNKCHeader wsNKC, False

    Application.ScreenUpdating = False

    ' Doc toan bo du lieu vao array 1 lan (cot A-J = 1-10)
    arr = wsNKC.Range(wsNKC.Cells(3, 1), wsNKC.Cells(lastRow, 10)).Value

    ' Xu ly trong memory
    For r = 1 To UBound(arr, 1)
        If IsError(arr(r, 1)) Then GoTo NextRowBSC
        If IsEmpty(arr(r, 1)) Or arr(r, 1) = "" Then GoTo NextRowBSC

            If IsError(arr(r, 4)) Then tkNo = "" Else tkNo = Trim$(CStr(arr(r, 4)))
            If IsError(arr(r, 5)) Then tkCo = "" Else tkCo = Trim$(CStr(arr(r, 5)))

            ' Bo sung Ngay hach toan (cot 7) = Ngay CT neu thieu
            If IsError(arr(r, 7)) Or IsEmpty(arr(r, 7)) Then
                arr(r, 7) = arr(r, 1)
            ElseIf arr(r, 7) = "" Or Not IsDate(arr(r, 7)) Then
                arr(r, 7) = arr(r, 1)
            End If
            ngayHT = arr(r, 7)

            ' Bo sung Thang (cot 8): uu tien Ngay hach toan, neu khong co thi lay Ngay chung tu
            Dim mVal As Variant, mCalc As Variant
            mVal = arr(r, 8)
            If IsError(mVal) Or IsEmpty(mVal) Or mVal = "" Or mVal = 0 Or Not IsNumeric(mVal) Or mVal < 1 Or mVal > 12 Then
                mCalc = GetMonthFromAnyDate(ngayHT, arr(r, 1))
                If IsEmpty(mCalc) Or mCalc = "" Then
                    arr(r, 8) = 0
                Else
                    arr(r, 8) = mCalc
                End If
            End If

            ' Cot I (No) va J (Co) la TK rut gon cap 3 cua cot D/E
            If tkNo <> "" Then arr(r, 9) = Left$(tkNo, 3)
            If tkCo <> "" Then arr(r, 10) = Left$(tkCo, 3)

            ' Cot F (So tien): neu bi blank/null thi set = 0
            amtVal = arr(r, 6)
            If IsError(amtVal) Or IsEmpty(amtVal) Then
                arr(r, 6) = 0
            ElseIf Len(Trim$(CStr(amtVal))) = 0 Then
                arr(r, 6) = 0
            End If
NextRowBSC:
    Next r

    ' Ghi lai 1 lan
    wsNKC.Range(wsNKC.Cells(3, 1), wsNKC.Cells(lastRow, 10)).Value = arr

    ' Tong tai E1/F1 bang SUBTOTAL (khong chen dong moi)
    wsNKC.Cells(1, 5).Value = "T" & ChrW(7893) & "ng :"
    wsNKC.Cells(1, 6).Formula = "=SUBTOTAL(9,F3:F" & lastRow & ")"
    wsNKC.Cells(1, 5).Font.Bold = True
    wsNKC.Cells(1, 6).Font.Bold = True
    wsNKC.Cells(1, 6).NumberFormat = "#,##0"
    wsNKC.Cells(1, 5).Font.Size = 11
    wsNKC.Cells(1, 6).Font.Size = 11

    Application.ScreenUpdating = True

    ' Dam bao nut Xoa loc luon co tren NKC
    FixClearFilterButton wsNKC
End Sub

' Đảm bảo header NKC có đủ cột "Khac" và "Can review" (khi mở file cũ)
Private Sub EnsureNKCHeader(ws As Worksheet, Optional includeReview As Boolean = False, Optional extraHeaders As Variant)

    ReorderNKCColumnsIfOld ws
    Const HDR_ROW As Long = 2

    Dim extraCount As Long, i As Long, lastCol As Long
    Dim headerVal As Variant

    extraCount = 0
    If Not IsMissing(extraHeaders) Then
        If IsArray(extraHeaders) Then
            On Error Resume Next
            extraCount = UBound(extraHeaders) - LBound(extraHeaders) + 1
            If extraCount < 0 Then extraCount = 0
            On Error GoTo 0
        End If
    End If

    ' Base headers (new order)
    ws.Cells(HDR_ROW, 1).Value = "Ng" & ChrW(224) & "y ch" & ChrW(7913) & "ng t" & ChrW(7915)
    ws.Cells(HDR_ROW, 2).Value = "S" & ChrW(7889) & " h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    ws.Cells(HDR_ROW, 3).Value = "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i"
    ws.Cells(HDR_ROW, 4).Value = "N" & ChrW(7907) & " TK"
    ws.Cells(HDR_ROW, 5).Value = "C" & ChrW(243) & " TK"
    ws.Cells(HDR_ROW, 6).Value = "S" & ChrW(7889) & " ti" & ChrW(7873) & "n"
    ws.Cells(HDR_ROW, 7).Value = "Ng" & ChrW(224) & "y ho" & ChrW(7841) & "ch to" & ChrW(225) & "n"
    ws.Cells(HDR_ROW, 8).Value = "Th" & ChrW(225) & "ng"
    ws.Cells(HDR_ROW, 9).Value = "N" & ChrW(7907)
    ws.Cells(HDR_ROW, 10).Value = "C" & ChrW(243)
    ws.Cells(HDR_ROW, 11).Value = "Kh" & ChrW(225) & "c"

    If extraCount > 0 Then
        For i = 1 To extraCount
            headerVal = extraHeaders(i)
            If IsError(headerVal) Or IsEmpty(headerVal) Then headerVal = "Khac " & i: GoTo NextHdr
            If Len(Trim$(CStr(headerVal))) = 0 Then headerVal = "Khac " & i
NextHdr:
            ws.Cells(HDR_ROW, 11 + i).Value = headerVal
        Next i
    End If

    If includeReview Then
        ws.Cells(HDR_ROW, 11 + extraCount + 1).Value = "C" & ChrW(7847) & "n review"
    End If

    lastCol = 11 + extraCount + IIf(includeReview, 1, 0)

    ws.Cells.Font.Name = "Times New Roman"
    With ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .AutoFilter
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).EntireColumn.AutoFit

    ' Limit width for Dien giai (C) to avoid overly wide columns
    If ws.Columns(3).ColumnWidth > 50 Then ws.Columns(3).ColumnWidth = 50
End Sub



Private Sub FillExtraFromPair(ByRef outputArr As Variant, ByVal outRow As Long, ByVal extraCount As Long, ByVal colExtraStart As Long, ByRef arrExtra As Variant, ByVal rNo As Long, ByVal rCo As Long)

    Dim j As Long, extraVal As Variant

    If extraCount <= 0 Then Exit Sub

    For j = 1 To extraCount

        extraVal = arrExtra(rNo, j)

        If IsError(extraVal) Or IsEmpty(extraVal) Then
            extraVal = arrExtra(rCo, j)
        ElseIf Len(Trim$(CStr(extraVal))) = 0 Then
            extraVal = arrExtra(rCo, j)
        End If

        outputArr(outRow, colExtraStart + j - 1) = extraVal

    Next j

End Sub



Private Sub FillExtraFromOne(ByRef outputArr As Variant, ByVal outRow As Long, ByVal extraCount As Long, ByVal colExtraStart As Long, ByRef arrExtra As Variant, ByVal rIdx As Long)

    Dim j As Long

    If extraCount <= 0 Then Exit Sub

    For j = 1 To extraCount

        outputArr(outRow, colExtraStart + j - 1) = arrExtra(rIdx, j)

    Next j

End Sub



Private Function NzDbl(ByVal v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzDbl = 0#
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0#
    End If
End Function

Private Function NzVal(ByVal v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzVal = 0#
    ElseIf IsNumeric(v) Then
        NzVal = CDbl(v)
    Else
        NzVal = CDbl(Val(CStr(v)))
    End If
End Function


Private Function GetLastUsedColumn(ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        GetLastUsedColumn = 0
    Else
        GetLastUsedColumn = lastCell.Column
    End If
End Function

Private Function GetLastUsedRow(ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = lastCell.Row
    End If
End Function

Private Function GetLastUsedRowInCol(ws As Worksheet, ByVal colIndex As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(colIndex).Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        GetLastUsedRowInCol = 0
    Else
        GetLastUsedRowInCol = lastCell.Row
    End If
End Function

Private Function AmountKey(ByVal v As Double) As String
    AmountKey = Format$(Round(v, 2), "0.00")
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
    wsTemplate.Name = "NKC"            ' Create header
    With wsTemplate
        .Cells(2, 1).Value = "Ng" & ChrW(224) & "y ch" & ChrW(7913) & "ng t" & ChrW(7915)
        .Cells(2, 2).Value = "S" & ChrW(7889) & " h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
        .Cells(2, 3).Value = "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i"
        .Cells(2, 4).Value = "N" & ChrW(7907) & " TK"
        .Cells(2, 5).Value = "C" & ChrW(243) & " TK"
        .Cells(2, 6).Value = "S" & ChrW(7889) & " ti" & ChrW(7873) & "n"
        .Cells(2, 7).Value = "Ng" & ChrW(224) & "y ho" & ChrW(7841) & "ch to" & ChrW(225) & "n"
        .Cells(2, 8).Value = "Th" & ChrW(225) & "ng"
        .Cells(2, 9).Value = "N" & ChrW(7907)
        .Cells(2, 10).Value = "C" & ChrW(243)
        .Cells(2, 11).Value = "Kh" & ChrW(225) & "c"
        .Cells.Font.Name = "Times New Roman"
        ' Format header (khong can cot review cho so da xu ly)
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
    If Not fastMode Then ApplyWorkbookFont wb, "Times New Roman"
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
    Dim lastRow As Long, i As Long, j As Long
    Dim arrData As Variant
    Dim rowCount As Long
    Dim arrMaCT() As String, arrNgay() As Variant, arrDienGiai() As Variant
    Dim arrTK() As String, arrTK3() As String
    Dim arrNo() As Double, arrCo() As Double
    Dim arrKhac() As Variant, arrMonth() As Variant
    Dim arrKey() As String
    Dim arrExtra() As Variant, extraHeaders() As Variant
    Dim extraCount As Long, lastColSrc As Long
    Dim colKhac As Long, colExtraStart As Long, colReview As Long
    Dim pairCache As Object
    Dim key As Variant, r As Variant
    Dim pivotErr As String, thMsg As String
    Dim wsNKCExists As Worksheet
    Dim isTemplateNKC As Boolean
    Dim includeReview As Boolean
    Dim maxGroupSize As Long
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldStatus As Variant
    Dim doHeavy As Boolean
    Const SAFE_MAX_GROUP_ROWS As Long = 2000
    Const FAST_FORCE_ROWS As Long = 120000
    Const FAST_AUTO_HEAVY As Boolean = True
    Const CHUNK_WRITE_ROWS As Long = 50000
    Const LEGACY_FAST_ROWS As Long = 50000
    Dim fastMode As Boolean
    Dim legacyMode As Boolean
    ' Mac dinh: xu ly du lieu tho -> co cot Can review
    includeReview = True

    Set wb = ActiveWorkbook
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar

    ' Check if NKC sheet already exists (user used template)
    On Error Resume Next
    Set wsNKCExists = wb.Sheets("NKC")
    On Error GoTo 0

    ' Check if "So Nhat Ky Chung" source sheet exists
    On Error Resume Next
    Set wsNguon = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0

    ' If no source sheet -> try to detect from ActiveSheet or other sheets
    If wsNguon Is Nothing Then
        Set wsNguon = FindSourceSheetSmart(wb, ActiveSheet)
        If wsNguon Is Nothing Then
            ' Treat as already processed if NKC has data
            If Not wsNKCExists Is Nothing And SheetHasDataRows(wsNKCExists, 3) Then
                InfoToast "Khong tim thay sheet nguon. Su dung NKC hien co."
                EnsureNKCHeader wsNKCExists, False
                includeReview = False
                doHeavy = True
                GoTo SkipProcessing
            End If
            MsgBox "Khong tim thay sheet 'So Nhat Ky Chung' hoac sheet nguon du lieu hop le.", vbExclamation
            Exit Sub
        End If
    End If

    ' If NKC exists AND source sheet exists -> Ask user what to do
    If Not wsNKCExists Is Nothing And Not wsNguon Is Nothing Then
        If Not ConfirmProceed("Sheet 'NKC' da ton tai. Xoa va tao lai tu 'So Nhat Ky Chung'?") Then
            InfoToast "Giu nguyen sheet NKC hien co. Bo qua xu ly."
            EnsureNKCHeader wsNKCExists, False
            includeReview = False
            doHeavy = True
            GoTo SkipProcessing
        End If
        ' User confirmed -> will rebuild NKC from source
    End If

    ' If source sheet exists but has no data, fall back to existing NKC if possible
    If Not wsNguon Is Nothing Then
        If Not SheetHasDataRows(wsNguon, 2) Then
            If Not wsNKCExists Is Nothing And SheetHasDataRows(wsNKCExists, 3) Then
                InfoToast "Sheet nguon khong co du lieu. Su dung NKC hien co."
                EnsureNKCHeader wsNKCExists, False
                includeReview = False
                doHeavy = True
                GoTo SkipProcessing
            Else
                MsgBox "Sheet nguon khong co du lieu.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    wsNguon.Activate
    Set wb = wsNguon.Parent
    lastRow = wsNguon.Cells(wsNguon.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Kh" & ChrW(244) & "ng c" & ChrW(243) & " d" & ChrW(7919) & " li" & ChrW(7879) & "u " & ChrW(273) & ChrW(7875) & " x" & ChrW(7917) & " l" & ChrW(253) & "!", vbExclamation
        Exit Sub
    End If
    ' Doc du lieu vao array truoc khi tao sheet moi
    lastColSrc = GetLastUsedColumn(wsNguon)
    If lastColSrc < 7 Then lastColSrc = 7
    arrData = wsNguon.Range(wsNguon.Cells(2, 1), wsNguon.Cells(lastRow, lastColSrc)).Value
    rowCount = UBound(arrData, 1)
    legacyMode = (rowCount <= LEGACY_FAST_ROWS)
    fastMode = (rowCount >= FAST_FORCE_ROWS)
    If Not legacyMode Then Application.StatusBar = "Dang doc du lieu NKC..."
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
    extraCount = lastColSrc - 7
    If extraCount > 0 Then
        ReDim arrExtra(1 To rowCount, 1 To extraCount)
        ReDim extraHeaders(1 To extraCount)
        For j = 1 To extraCount
            If IsError(wsNguon.Cells(1, 7 + j).Value) Then
                extraHeaders(j) = "Khac " & j
            Else
                extraHeaders(j) = Trim$(CStr(wsNguon.Cells(1, 7 + j).Value))
                If extraHeaders(j) = "" Then extraHeaders(j) = "Khac " & j
            End If
        Next j
    End If
    If Not legacyMode Then Application.StatusBar = "Dang xu ly du lieu (" & rowCount & " dong)..."
    For i = 1 To rowCount
        If IsError(arrData(i, 1)) Then arrMaCT(i) = "" Else arrMaCT(i) = Trim$(CStr(arrData(i, 1)))
        arrNgay(i) = arrData(i, 2)
        arrDienGiai(i) = arrData(i, 3)
        If IsError(arrData(i, 4)) Then arrTK(i) = "" Else arrTK(i) = Trim$(CStr(arrData(i, 4)))
        arrNo(i) = NzDbl(arrData(i, 5))
        arrCo(i) = NzDbl(arrData(i, 6))
        arrKhac(i) = arrData(i, 7)
        If extraCount > 0 Then
            For j = 1 To extraCount
                arrExtra(i, j) = arrData(i, 7 + j)
            Next j
        End If
        If arrTK(i) <> "" Then
            arrTK3(i) = Left$(arrTK(i), 3)
        Else
            arrTK3(i) = ""
        End If
        arrMonth(i) = GetMonthValue(arrNgay(i))
        If arrMaCT(i) <> "" Then
            If IsError(arrNgay(i)) Then
                arrKey(i) = arrMaCT(i) & "|"
            Else
                arrKey(i) = arrMaCT(i) & "|" & Trim$(CStr(arrNgay(i)))
            End If
        End If
        If (i Mod 5000) = 0 Then
            If Not legacyMode Then
                Application.StatusBar = "Dang xu ly du lieu (" & i & "/" & rowCount & ")..."
                DoEvents
            End If
        End If
    Next i
    ' Neu khong co nhom lon, uu tien duong xu ly cu cho nhanh
    If maxGroupSize <= SAFE_MAX_GROUP_ROWS Then
        legacyMode = True
        fastMode = False
    End If
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

    .Cells(2, 1).Value = "Ng" & ChrW(224) & "y ch" & ChrW(7913) & "ng t" & ChrW(7915)
    .Cells(2, 2).Value = "S" & ChrW(7889) & " h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    .Cells(2, 3).Value = "Di" & ChrW(7877) & "n gi" & ChrW(7843) & "i"
    .Cells(2, 4).Value = "N" & ChrW(7907) & " TK"
    .Cells(2, 5).Value = "C" & ChrW(243) & " TK"
    .Cells(2, 6).Value = "S" & ChrW(7889) & " ti" & ChrW(7873) & "n"
    .Cells(2, 7).Value = "Ng" & ChrW(224) & "y ho" & ChrW(7841) & "ch to" & ChrW(225) & "n"
    .Cells(2, 8).Value = "Th" & ChrW(225) & "ng"
    .Cells(2, 9).Value = "N" & ChrW(7907)
    .Cells(2, 10).Value = "C" & ChrW(243)

End With

' Chuan hoa header (Khac + review + cot bo sung)

EnsureNKCHeader wsKetQua, True, extraHeaders
    Set dictGroup = CreateObject("Scripting.Dictionary")
    ' Nhom du lieu theo MaCT|Ngay
    If Not legacyMode Then Application.StatusBar = "Dang nhom du lieu..."
    For i = 1 To rowCount
        If arrMaCT(i) <> "" Then
            key = arrKey(i)
            If Not dictGroup.Exists(key) Then dictGroup.Add key, New Collection
            dictGroup(key).Add i
            If dictGroup(key).Count > maxGroupSize Then maxGroupSize = dictGroup(key).Count
        End If
        If (i Mod 5000) = 0 Then
            If Not legacyMode Then
                Application.StatusBar = "Dang nhom du lieu (" & i & "/" & rowCount & ")..."
                DoEvents
            End If
        End If
    Next i
    ' ========== BUOC 1: XAC DINH NHOM "BAN" ==========
    ' Nhom "ban" = co it nhat 1 dong co CA No va Co
    Dim dictDirty As Object
    Set dictDirty = CreateObject("Scripting.Dictionary")
    Set pairCache = CreateObject("Scripting.Dictionary")
    Dim groupIndex As Long, groupTotal As Long
    groupIndex = 0
    groupTotal = dictGroup.Count
    For Each key In dictGroup.keys
        groupIndex = groupIndex + 1
        If (groupIndex Mod 50) = 0 Then
            If Not legacyMode Then
                Application.StatusBar = "Dang xu ly nhom " & groupIndex & "/" & groupTotal & "..."
                DoEvents
            End If
        End If
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
    colKhac = 11

    colExtraStart = colKhac + 1

    colReview = 0

    colCount = colKhac + extraCount + IIf(includeReview, 1, 0)

    If includeReview Then colReview = colKhac + extraCount + 1
    Dim initialCap As Long
    initialCap = rowCount * 2
    If initialCap < 1000 Then initialCap = 1000
    ReDim outputArr(1 To initialCap, 1 To colCount)
    Dim dongOut As Long
    dongOut = 1
    For Each key In dictGroup.keys
        Dim dsNoEntries As Collection, dsCoEntries As Collection
        Set dsNoEntries = New Collection
        Set dsCoEntries = New Collection
        Dim groupSize As Long
        groupSize = dictGroup(key).Count
        If groupSize > SAFE_MAX_GROUP_ROWS Then
            fastMode = True
            Application.StatusBar = "Nhom qua lon (" & groupSize & " dong) - xuat nhanh can review..."
            Dim cntFast As Long
            cntFast = 0
            For Each r In dictGroup(key)
                If arrNo(r) <> 0 Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    outputArr(dongOut, 1) = arrNgay(r)
                    outputArr(dongOut, 2) = arrMaCT(r)
                    outputArr(dongOut, 3) = arrDienGiai(r)
                    outputArr(dongOut, 4) = arrTK(r)
                    outputArr(dongOut, 5) = ""
                    outputArr(dongOut, 6) = arrNo(r)
                    outputArr(dongOut, 7) = arrNgay(r)
                    outputArr(dongOut, 8) = arrMonth(r)
                    outputArr(dongOut, 9) = arrTK3(r)
                    outputArr(dongOut, 10) = ""
                    outputArr(dongOut, colKhac) = arrKhac(r)
                    FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, r
                    If includeReview Then outputArr(dongOut, colReview) = "X"
                    dongOut = dongOut + 1
                End If
                If arrCo(r) <> 0 Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    outputArr(dongOut, 1) = arrNgay(r)
                    outputArr(dongOut, 2) = arrMaCT(r)
                    outputArr(dongOut, 3) = arrDienGiai(r)
                    outputArr(dongOut, 4) = ""
                    outputArr(dongOut, 5) = arrTK(r)
                    outputArr(dongOut, 6) = arrCo(r)
                    outputArr(dongOut, 7) = arrNgay(r)
                    outputArr(dongOut, 8) = arrMonth(r)
                    outputArr(dongOut, 9) = ""
                    outputArr(dongOut, 10) = arrTK3(r)
                    outputArr(dongOut, colKhac) = arrKhac(r)
                    FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, r
                    If includeReview Then outputArr(dongOut, colReview) = "X"
                    dongOut = dongOut + 1
                End If
                cntFast = cntFast + 1
                If (cntFast Mod 500) = 0 Then DoEvents
            Next r
            GoTo NextGroup
        End If
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
        Dim usedNo() As Double, usedCo() As Double
        Dim idxNo As Long, idxCo As Long
        Dim entryNo As Variant, entryCo As Variant
        Dim rNo As Long, rCo As Long
        Dim tienNoEntry As Double, tienCoEntry As Double
        Dim tienNo As Double, tienCo As Double
        Dim tienPhanBo As Double
        Dim absNo As Double, absCo As Double
        Dim tkNo As String, tkCo As String
        Dim khacValFast As Variant, extraVal As Variant
        Dim totalNo As Double, totalCo As Double
        Dim canFastPath As Boolean
        Dim allowCrossSign As Boolean, crossSignExactOnly As Boolean
        Dim mapCo As Object, keyAmt As String, keyOpp As String
        Dim matchedCount As Long, idxCoMatch As Long
        Dim coll As Collection
        Dim skipDeepPass As Boolean
        ' Neu can xu ly truong hop co ca am/duong trong 1 ma CT
        ' allowCrossSign = True: cho phep ghep cheo dau (am/duong)
        ' crossSignExactOnly = True: chi cho phep ghep cheo dau khi so tien khop 1-1
        allowCrossSign = True
        crossSignExactOnly = True
        ' Lay trang thai "ban" cua nhom
        skipDeepPass = (groupSize > SAFE_MAX_GROUP_ROWS)
        If dsNoEntries.Count = 0 Or dsCoEntries.Count = 0 Then
            ' Output single-side entries for review
            For idxNo = 1 To dsNoEntries.Count
                entryNo = dsNoEntries(idxNo)
                rNo = entryNo(0)
                tienNoEntry = entryNo(1)
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                outputArr(dongOut, 1) = arrNgay(rNo)        ' Ngay chung tu
                outputArr(dongOut, 2) = arrMaCT(rNo)       ' So CT
                outputArr(dongOut, 3) = arrDienGiai(rNo)   ' Dien giai
                outputArr(dongOut, 4) = arrTK(rNo)         ' TK No (full)
                outputArr(dongOut, 5) = ""                ' TK Co (full)
                outputArr(dongOut, 6) = tienNoEntry        ' So tien
                outputArr(dongOut, 7) = arrNgay(rNo)       ' Ngay hach toan
                outputArr(dongOut, 8) = arrMonth(rNo)      ' Thang
                outputArr(dongOut, 9) = arrTK3(rNo)        ' No (3 ky tu)
                outputArr(dongOut, 10) = ""               ' Co (3 ky tu)
                outputArr(dongOut, colKhac) = arrKhac(rNo)
                FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo
                If includeReview Then outputArr(dongOut, colReview) = "X"
                dongOut = dongOut + 1
            Next idxNo
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                outputArr(dongOut, 1) = arrNgay(rCo)        ' Ngay chung tu
                outputArr(dongOut, 2) = arrMaCT(rCo)       ' So CT
                outputArr(dongOut, 3) = arrDienGiai(rCo)   ' Dien giai
                outputArr(dongOut, 4) = ""                ' TK No (full)
                outputArr(dongOut, 5) = arrTK(rCo)         ' TK Co (full)
                outputArr(dongOut, 6) = tienCoEntry        ' So tien
                outputArr(dongOut, 7) = arrNgay(rCo)       ' Ngay hach toan
                outputArr(dongOut, 8) = arrMonth(rCo)      ' Thang
                outputArr(dongOut, 9) = ""                ' No (3 ky tu)
                outputArr(dongOut, 10) = arrTK3(rCo)       ' Co (3 ky tu)
                outputArr(dongOut, colKhac) = arrKhac(rCo)
                FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, rCo
                If includeReview Then outputArr(dongOut, colReview) = "X"
                dongOut = dongOut + 1
            Next idxCo
            GoTo NextGroup
        End If
        ReDim usedNo(1 To dsNoEntries.Count)
        ReDim usedCo(1 To dsCoEntries.Count)
        totalNo = 0#
        For idxNo = 1 To dsNoEntries.Count
            totalNo = totalNo + dsNoEntries(idxNo)(1)
        Next idxNo
        totalCo = 0#
        For idxCo = 1 To dsCoEntries.Count
            totalCo = totalCo + dsCoEntries(idxCo)(1)
        Next idxCo
        canFastPath = (Abs(totalNo - totalCo) < 0.01)
        ' ========== FAST PATH: 1 NO or 1 CO -> phan bo truc tiep (giu dung gia tri am) ==========
        If canFastPath And (dsNoEntries.Count = 1 Or dsCoEntries.Count = 1) Then
            Dim signNoFast As Integer, signCoFast As Integer, fastOk As Boolean
            fastOk = True
            If dsNoEntries.Count = 1 Then
                signNoFast = Sgn(dsNoEntries(1)(1))
                For idxCo = 1 To dsCoEntries.Count
                    signCoFast = Sgn(dsCoEntries(idxCo)(1))
                    If signCoFast <> 0 And signNoFast <> 0 And signCoFast <> signNoFast Then
                        If (Not allowCrossSign) Or crossSignExactOnly Then fastOk = False
                        Exit For
                    End If
                Next idxCo
            Else
                signCoFast = Sgn(dsCoEntries(1)(1))
                For idxNo = 1 To dsNoEntries.Count
                    signNoFast = Sgn(dsNoEntries(idxNo)(1))
                    If signNoFast <> 0 And signCoFast <> 0 And signNoFast <> signCoFast Then
                        If (Not allowCrossSign) Or crossSignExactOnly Then fastOk = False
                        Exit For
                    End If
                Next idxNo
            End If
            If Not fastOk Then GoTo SkipFastPath
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
                    If IsError(khacValFast) Or IsEmpty(khacValFast) Then khacValFast = arrKhac(rCo) _
                    Else: If Len(Trim$(CStr(khacValFast))) = 0 Then khacValFast = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                    outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                    outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                    outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                    outputArr(dongOut, 6) = tienCoEntry     ' So tien (giu dung dau)  ' So tien
                    outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                    outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                    outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                    outputArr(dongOut, colKhac) = khacValFast     ' Khac
                    FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                    If includeReview Then
                        If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                            outputArr(dongOut, colReview) = ""
                        Else
                            outputArr(dongOut, colReview) = "X"
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
                    If IsError(khacValFast) Or IsEmpty(khacValFast) Then khacValFast = arrKhac(rCo) _
                    Else: If Len(Trim$(CStr(khacValFast))) = 0 Then khacValFast = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                    outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                    outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                    outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                    outputArr(dongOut, 6) = tienNoEntry     ' So tien (giu dung dau)  ' So tien
                    outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                    outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                    outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                    outputArr(dongOut, colKhac) = khacValFast     ' Khac

                    FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                    If includeReview Then
                        If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                            outputArr(dongOut, colReview) = ""
                        Else
                            outputArr(dongOut, colReview) = "X"
                        End If
                    End If
                    usedNo(idxNo) = usedNo(idxNo) + tienNoEntry
                    usedCo(1) = usedCo(1) + tienNoEntry
                    dongOut = dongOut + 1
                Next idxNo
            End If
            GoTo NextGroup
        End If
SkipFastPath:
        If Not legacyMode Then
            ' ========== QUICK MATCH BY AMOUNT (hash) ==========
            ' Ghep nhanh theo so tien trung khop de giam nhom lon
            Set mapCo = CreateObject("Scripting.Dictionary")
            For idxCo = 1 To dsCoEntries.Count
                keyAmt = AmountKey(dsCoEntries(idxCo)(1))
                If Not mapCo.Exists(keyAmt) Then
                    Set coll = New Collection
                    mapCo.Add keyAmt, coll
                End If
                mapCo(keyAmt).Add idxCo
            Next idxCo
            matchedCount = 0
            For idxNo = 1 To dsNoEntries.Count
                entryNo = dsNoEntries(idxNo)
                rNo = entryNo(0)
                tienNoEntry = entryNo(1)
                keyAmt = AmountKey(tienNoEntry)
                idxCoMatch = 0
                If mapCo.Exists(keyAmt) Then
                    Set coll = mapCo(keyAmt)
                    If coll.Count > 0 Then
                        idxCoMatch = coll(1)
                        coll.Remove 1
                    End If
                ElseIf allowCrossSign And crossSignExactOnly Then
                    keyOpp = AmountKey(-tienNoEntry)
                    If mapCo.Exists(keyOpp) Then
                        Set coll = mapCo(keyOpp)
                        If coll.Count > 0 Then
                            idxCoMatch = coll(1)
                            coll.Remove 1
                        End If
                    End If
                End If
                If idxCoMatch > 0 Then
                    entryCo = dsCoEntries(idxCoMatch)
                    rCo = entryCo(0)
                    tkNo = arrTK(rNo)
                    tkCo = arrTK(rCo)
                    Dim dgNoQuick As String, dgCoQuick As String
                    Dim sameDGQuick As Boolean, samePrefixQuick As Boolean
                    dgNoQuick = NormalizeDG(arrDienGiai(rNo))
                    dgCoQuick = NormalizeDG(arrDienGiai(rCo))
                    sameDGQuick = (dgNoQuick = dgCoQuick)
                    samePrefixQuick = (Left$(tkNo, 3) = Left$(tkCo, 3))
                    If Not (sameDGQuick Or samePrefixQuick) Then GoTo NextNoQuick
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    khacValFast = arrKhac(rNo)
                    If IsError(khacValFast) Or IsEmpty(khacValFast) Then khacValFast = arrKhac(rCo) _
                    Else: If Len(Trim$(CStr(khacValFast))) = 0 Then khacValFast = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                    outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                    outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                    outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                    outputArr(dongOut, 6) = tienNoEntry     ' So tien (giu dung dau)
                    outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                    outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                    outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                    outputArr(dongOut, colKhac) = khacValFast     ' Khac
                    FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                    If includeReview Then
                        If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                            outputArr(dongOut, colReview) = ""
                        Else
                            outputArr(dongOut, colReview) = "X"
                        End If
                    End If
                    usedNo(idxNo) = usedNo(idxNo) + tienNoEntry
                    usedCo(idxCoMatch) = usedCo(idxCoMatch) + tienNoEntry
                    dongOut = dongOut + 1
                    matchedCount = matchedCount + 1
                End If
NextNoQuick:
                If (matchedCount Mod 500) = 0 Then If Not legacyMode Then DoEvents
            Next idxNo
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
                If Sgn(tienNo) <> 0 And Sgn(tienCo) <> 0 Then
                    If Sgn(tienNo) <> Sgn(tienCo) Then
                        If Not allowCrossSign Then GoTo NextCoPass1
                    End If
                End If
                tkCo = arrTK(rCo)
                If Len(tkNo) > 3 And Len(tkCo) > 3 Then
                    If tkNo = tkCo Then GoTo NextCoPass1
                End If
                If Abs(tienNo - tienCo) < 0.01 And IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                    End If
                    ' Format mau: NgayHT, NgayCT, Thang, SoHD, DienGiai, No, Co, NoTK, CoTK, SoTien, Khac, CanReview
                    Dim khacVal As Variant
                    khacVal = arrKhac(rNo)
                    If IsError(khacVal) Or IsEmpty(khacVal) Then khacVal = arrKhac(rCo) _
                    Else: If Len(Trim$(CStr(khacVal))) = 0 Then khacVal = arrKhac(rCo)
                    outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                    outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                    outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                    outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                    outputArr(dongOut, 6) = tienNo          ' So tien  ' So tien
                    outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                    outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                    outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                    outputArr(dongOut, colKhac) = khacVal          ' Khac (lay tu G, uu tien dong No, neu trong thi dong Co)
                    FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                    usedNo(idxNo) = usedNo(idxNo) + tienNo
                    usedCo(idxCo) = usedCo(idxCo) + tienNo
                    dongOut = dongOut + 1
                    Exit For
                End If
NextCoPass1:
            Next idxCo
NextNoPass1:
        Next idxNo
        If Not skipDeepPass Then
        ' ========== PASS 2: Ghep so tien khop (uu tien cung dien giai neu co) ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass2
            Dim dgNo2 As String
            dgNo2 = NormalizeDG(arrDienGiai(rNo))
            Dim hasValidCandidate As Boolean
            hasValidCandidate = False
            tkNo = arrTK(rNo)
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoCheck2
                If Sgn(tienNo) <> 0 And Sgn(tienCo) <> 0 Then
                    If Sgn(tienNo) <> Sgn(tienCo) Then
                        If Not allowCrossSign Then GoTo NextCoCheck2
                    End If
                End If
                If Abs(tienNo - tienCo) < 0.01 Then
                    tkCo = arrTK(rCo)
                    If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                        hasValidCandidate = True
                        Exit For
                    End If
                End If
NextCoCheck2:
            Next idxCo
            Dim bestIdxPass2 As Long
            Dim bestScore As Long
            bestIdxPass2 = 0
            bestScore = -999
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoPass2
                If Sgn(tienNo) <> 0 And Sgn(tienCo) <> 0 Then
                    If Sgn(tienNo) <> Sgn(tienCo) Then
                        If Not allowCrossSign Then GoTo NextCoPass2
                    End If
                End If
                If Abs(tienNo - tienCo) < 0.01 Then
                    Dim score As Long
                    Dim sameDG As Boolean
                    Dim isValid As Boolean
                    Dim sameTK As Boolean
                    sameDG = (dgNo2 = NormalizeDG(arrDienGiai(rCo)))
                    isValid = IsValidAccountPairCached(tkNo, arrTK(rCo), pairCache)
                    If hasValidCandidate And Not isValid Then GoTo NextCoPass2
                    sameTK = (Left$(arrTK(rNo), 3) = Left$(arrTK(rCo), 3))
                    score = 0
                    If isValid Then score = score + 10
                    If sameDG Then score = score + 5
                    If Not sameTK Then score = score + 1 Else score = score - 100
                    If score > bestScore Then
                        bestScore = score
                        bestIdxPass2 = idxCo
                    End If
                End If
NextCoPass2:
            Next idxCo
            If bestIdxPass2 > 0 Then
                entryCo = dsCoEntries(bestIdxPass2)
                rCo = entryCo(0)
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                Dim khacVal2 As Variant
                khacVal2 = arrKhac(rNo)
                If IsError(khacVal2) Or IsEmpty(khacVal2) Then khacVal2 = arrKhac(rCo) _
                Else: If Len(Trim$(CStr(khacVal2))) = 0 Then khacVal2 = arrKhac(rCo)
                outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                outputArr(dongOut, 6) = tienNo          ' So tien
                outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                outputArr(dongOut, colKhac) = khacVal2        ' Khac
                FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                usedNo(idxNo) = usedNo(idxNo) + tienNo
                usedCo(bestIdxPass2) = usedCo(bestIdxPass2) + tienNo
                dongOut = dongOut + 1
            End If
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
                        If Sgn(tienNo) <> 0 And Sgn(tienCo) <> 0 Then
                            If Sgn(tienNo) <> Sgn(tienCo) Then
                                If (Not allowCrossSign) Or crossSignExactOnly Then GoTo NextCoPick
                            End If
                        End If
                        If bestIdx = 0 Or Abs(tienCo) < bestRemain Then
                            bestIdx = idxCo
                            bestRemain = Abs(tienCo)
                            bestCo = tienCo
                        End If
                    End If
NextCoPick:
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
                If IsError(khacVal3) Or IsEmpty(khacVal3) Then khacVal3 = arrKhac(rCo) _
                Else: If Len(Trim$(CStr(khacVal3))) = 0 Then khacVal3 = arrKhac(rCo)
                outputArr(dongOut, 1) = arrNgay(rNo)  ' Ngay chung tu
                    outputArr(dongOut, 2) = arrMaCT(rNo)  ' So CT
                    outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                    outputArr(dongOut, 4) = arrTK(rNo)  ' TK No (full)
                    outputArr(dongOut, 5) = arrTK(rCo)  ' TK Co (full)
                    outputArr(dongOut, 6) = tienPhanBo      ' So tien  ' So tien
                    outputArr(dongOut, 7) = arrNgay(rNo)  ' Ngay hach toan
                    outputArr(dongOut, 8) = arrMonth(rNo)  ' Thang
                    outputArr(dongOut, 9) = arrTK3(rNo)  ' No (3 ky tu)
                    outputArr(dongOut, 10) = arrTK3(rCo)  ' Co (3 ky tu)
                outputArr(dongOut, colKhac) = khacVal3        ' Khac
                FillExtraFromPair outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo, rCo
                ' Nếu luồng chưa xử lý (includeReview=True) thì cột 12 là review
                If includeReview Then
                    If IsValidAccountPairCached(tkNo, tkCo, pairCache) Then
                        outputArr(dongOut, colReview) = ""
                    Else
                        outputArr(dongOut, colReview) = "X"
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
        End If
        ' ========== LEFTOVER: output unmatched amounts for review ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) >= 0.01 Then
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                outputArr(dongOut, 1) = arrNgay(rNo)      ' Ngay chung tu
                outputArr(dongOut, 2) = arrMaCT(rNo)      ' So CT
                outputArr(dongOut, 3) = arrDienGiai(rNo)  ' Dien giai
                outputArr(dongOut, 4) = arrTK(rNo)        ' TK No (full)
                outputArr(dongOut, 5) = ""               ' TK Co (blank)
                outputArr(dongOut, 6) = tienNo           ' So tien
                outputArr(dongOut, 7) = arrNgay(rNo)      ' Ngay hach toan
                outputArr(dongOut, 8) = arrMonth(rNo)     ' Thang
                outputArr(dongOut, 9) = arrTK3(rNo)       ' No (3 ky tu)
                outputArr(dongOut, 10) = ""              ' Co (blank)
                outputArr(dongOut, colKhac) = arrKhac(rNo)

                FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, rNo
                If includeReview Then outputArr(dongOut, colReview) = "X"
                dongOut = dongOut + 1
            End If
        Next idxNo
        For idxCo = 1 To dsCoEntries.Count
            entryCo = dsCoEntries(idxCo)
            rCo = entryCo(0)
            tienCoEntry = entryCo(1)
            tienCo = tienCoEntry - usedCo(idxCo)
            If Abs(tienCo) >= 0.01 Then
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To colCount)
                End If
                outputArr(dongOut, 1) = arrNgay(rCo)      ' Ngay chung tu
                outputArr(dongOut, 2) = arrMaCT(rCo)      ' So CT
                outputArr(dongOut, 3) = arrDienGiai(rCo)  ' Dien giai
                outputArr(dongOut, 4) = ""               ' TK No (blank)
                outputArr(dongOut, 5) = arrTK(rCo)        ' TK Co (full)
                outputArr(dongOut, 6) = tienCo           ' So tien
                outputArr(dongOut, 7) = arrNgay(rCo)      ' Ngay hach toan
                outputArr(dongOut, 8) = arrMonth(rCo)     ' Thang
                outputArr(dongOut, 9) = ""               ' No (blank)
                outputArr(dongOut, 10) = arrTK3(rCo)      ' Co (3 ky tu)
                outputArr(dongOut, colKhac) = arrKhac(rCo)

                FillExtraFromOne outputArr, dongOut, extraCount, colExtraStart, arrExtra, rCo
                If includeReview Then outputArr(dongOut, colReview) = "X"
                dongOut = dongOut + 1
            End If
        Next idxCo
NextGroup:
    Next key
    ' ========== GHI OUTPUT ==========
    Dim finalOut() As Variant
    If dongOut > 1 Then
        If Not legacyMode Then Application.StatusBar = "Dang ghi ket qua NKC..."
        If legacyMode Then
            ReDim finalOut(1 To dongOut - 1, 1 To colCount)
            For i = 1 To dongOut - 1
                For j = 1 To colCount
                    finalOut(i, j) = outputArr(i, j)
                Next j
            Next i
            wsKetQua.Range("A3").Resize(dongOut - 1, colCount).Value = finalOut
        ElseIf (Not fastMode) And (dongOut - 1 < CHUNK_WRITE_ROWS) Then
            ReDim finalOut(1 To dongOut - 1, 1 To colCount)
            For i = 1 To dongOut - 1
                For j = 1 To colCount
                    finalOut(i, j) = outputArr(i, j)
                Next j
            Next i
            wsKetQua.Range("A3").Resize(dongOut - 1, colCount).Value = finalOut
        Else
            Dim chunkSize As Long
            Dim startRow As Long, endRow As Long, chunkRows As Long
            Dim tempOut() As Variant
            chunkSize = 5000
            startRow = 1
            Do While startRow <= dongOut - 1
                endRow = startRow + chunkSize - 1
                If endRow > dongOut - 1 Then endRow = dongOut - 1
                chunkRows = endRow - startRow + 1
                ReDim tempOut(1 To chunkRows, 1 To colCount)
                For i = 1 To chunkRows
                    For j = 1 To colCount
                        tempOut(i, j) = outputArr(startRow + i - 1, j)
                    Next j
                Next i
                wsKetQua.Range("A" & (startRow + 2)).Resize(chunkRows, colCount).Value = tempOut
                startRow = endRow + 1
                DoEvents
            Loop
        End If
        EnsureNKCHeader wsKetQua, includeReview, extraHeaders
        wsKetQua.Cells.Font.Name = "Times New Roman"
        FixClearFilterButton wsKetQua
        ' ========== TO VANG CAC DONG CAN REVIEW ==========
        ' Tô vàng cột review (chỉ khi includeReview=True) - dung Union() de to 1 lan
        Dim rng As Range, rngReview As Range
        If includeReview And Not fastMode And Not legacyMode Then
            For i = 3 To dongOut + 1
                If wsKetQua.Cells(i, colReview).Value = "X" Then
                    Set rng = wsKetQua.Range(wsKetQua.Cells(i, 1), wsKetQua.Cells(i, colReview))
                    If rngReview Is Nothing Then
                        Set rngReview = rng
                    Else
                        Set rngReview = Union(rngReview, rng)
                    End If
                End If
                If (i Mod 5000) = 0 Then If Not legacyMode Then DoEvents
            Next i
            If Not rngReview Is Nothing Then rngReview.Interior.Color = RGB(255, 255, 150)
        End If
    End If
    ' ========== FORMAT ==========
    Dim lastRowOut As Long
    lastRowOut = dongOut + 1
    wsKetQua.Cells(1, 6).Formula = "=SUBTOTAL(9,F3:F" & lastRowOut & ")"
    wsKetQua.Cells(1, 6).Font.Bold = True
    wsKetQua.Cells(1, 5).Value = "Tong:"
    wsKetQua.Cells(1, 5).Font.Bold = True
    wsKetQua.Columns("F").NumberFormat = "#,##0"
    wsKetQua.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    wsKetQua.Columns("G:G").NumberFormat = "dd/mm/yyyy"
    ApplyWorkbookFont wb, "Times New Roman"
    ' Dem so dong can review
    Dim countReview As Long
    If includeReview And Not fastMode Then
        countReview = Application.WorksheetFunction.CountIf(wsKetQua.Columns(colReview), "X")
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
    doHeavy = (Not fastMode) Or FAST_AUTO_HEAVY
    If fastMode And Not doHeavy Then
        WarnToast "Fast mode: Bo qua tinh TB/TH/Pivot de tranh treo."
    End If

    If doHeavy Then
        If WorksheetExists("TB", wb) Then
            tbMsg = Auto_Tinh_TB(wsKetQua)
            If tbMsg <> "" Then InfoToast "T" & ChrW(205) & "NH TO" & ChrW(193) & "N TB TH" & ChrW(192) & "NH C" & ChrW(212) & "NG! " & tbMsg
        Else
            WarnToast "Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y sheet TB! T" & ChrW(7841) & "o m" & ChrW(7851) & "u TB tr" & ChrW(432) & ChrW(7899) & "c r" & ChrW(7891) & "i t" & ChrW(237) & "nh."
        End If
    End If
    If doHeavy Then
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
    End If

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

    If doHeavy Then
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
    End If

    ' Additional steps after processing (or skipping)
    MsgBox ChrW(272) & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t!" & vbCrLf & _
           "C" & ChrW(243) & " th" & ChrW(7875) & " ch" & ChrW(7841) & "y ti" & ChrW(7871) & "p c" & ChrW(225) & "c b" & ChrW(432) & ChrW(7899) & "c kh" & ChrW(225) & "c n" & ChrW(7871) & "u c" & ChrW(7847) & "n.", vbInformation
    Application.StatusBar = False
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
    Dim useLeftMatch As Boolean
    Dim warnNonNum As String
    Dim oldCalc As XlCalculation
    Dim arrNKC As Variant, arrSource As Variant
    Dim dictNoPrefix As Object, dictCoPrefix As Object
    Dim dictNoExact As Object, dictCoExact As Object
    Dim neededLens As Object
    Dim tkNoNKC As String, tkCoNKC As String, stien As Double
    Dim pfx As String, pLen As Long, lenKey As Variant
    Dim sumNo As Double, sumCo As Double
    Dim rngYellow As Range, rngGreen As Range, rngBold As Range, rngClear As Range
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
    ' Dam bao cot so E:J cua TB la so
    warnNonNum = NormalizeTBNumberColumns(wsTB, 4, lastRowTB)
    If Len(warnNonNum) > 0 Then WarnToast warnNonNum
    ' Tao dictionary chua danh sach TK tu sheet "So Nhat Ky Chung" (array-based)
    Set dictSourceTK = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsSource = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo ErrorHandler
    If Not wsSource Is Nothing Then
        lastRowSource = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
        If lastRowSource >= 2 Then
            arrSource = wsSource.Range("D2:D" & lastRowSource).Value
            For i = 1 To UBound(arrSource, 1)
                If IsError(arrSource(i, 1)) Or IsEmpty(arrSource(i, 1)) Then GoTo NextSrc
                tkFull = Trim$(CStr(arrSource(i, 1)))
                If tkFull <> "" Then dictSourceTK(tkFull) = True
NextSrc:
            Next i
        End If
    End If
    Application.Calculation = xlCalculationManual
    ' ========== DOC NKC VAO ARRAY 1 LAN ==========
    arrNKC = wsNKC.Range("D3:F" & lastRowNKC).Value  ' D=TK No, E=TK Co, F=So tien
    ' Thu thap cac do dai TK can tinh tu TB
    Set neededLens = CreateObject("Scripting.Dictionary")
    neededLens(3) = True
    For r = 4 To lastRowTB
        If IsError(wsTB.Cells(r, 3).Value) Then GoTo NextLen
        tkTrim = Trim$(CStr(wsTB.Cells(r, 3).Value))
        If tkTrim <> "" And Len(tkTrim) >= 3 Then neededLens(Len(tkTrim)) = True
NextLen:
    Next r
    ' Tao dictionary tong theo prefix va exact
    Set dictNoPrefix = CreateObject("Scripting.Dictionary")
    Set dictCoPrefix = CreateObject("Scripting.Dictionary")
    Set dictNoExact = CreateObject("Scripting.Dictionary")
    Set dictCoExact = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(arrNKC, 1)
        If IsError(arrNKC(r, 1)) Then tkNoNKC = "" Else tkNoNKC = Trim$(CStr(arrNKC(r, 1)))
        If IsError(arrNKC(r, 2)) Then tkCoNKC = "" Else tkCoNKC = Trim$(CStr(arrNKC(r, 2)))
        If IsNumeric(arrNKC(r, 3)) Then stien = CDbl(arrNKC(r, 3)) Else stien = 0#
        ' Exact match
        If tkNoNKC <> "" Then dictNoExact(tkNoNKC) = CDbl(dictNoExact(tkNoNKC)) + stien
        If tkCoNKC <> "" Then dictCoExact(tkCoNKC) = CDbl(dictCoExact(tkCoNKC)) + stien
        ' Prefix match cho moi do dai can thiet
        For Each lenKey In neededLens.keys
            pLen = CLng(lenKey)
            If Len(tkNoNKC) >= pLen Then
                pfx = CStr(pLen) & "|" & Left$(tkNoNKC, pLen)
                dictNoPrefix(pfx) = CDbl(dictNoPrefix(pfx)) + stien
            End If
            If Len(tkCoNKC) >= pLen Then
                pfx = CStr(pLen) & "|" & Left$(tkCoNKC, pLen)
                dictCoPrefix(pfx) = CDbl(dictCoPrefix(pfx)) + stien
            End If
        Next lenKey
    Next r
    ' ========== TINH TOAN TB ==========
    wsTB.Columns(3).NumberFormat = "@"
    For r = 4 To lastRowTB
        If IsError(wsTB.Cells(r, 3).Value) Then GoTo NextTBCalc
        If wsTB.Cells(r, 3).Value <> "" Then
            tkFull = CStr(wsTB.Cells(r, 3).Value)
            tkTrim = Trim$(tkFull)
            wsTB.Cells(r, 3).Value = "'" & tkTrim
            wsTB.Cells(r, 2).Value = Len(tkTrim)
            useLeftMatch = (Len(tkTrim) >= 6 And Not dictSourceTK.Exists(tkTrim))
            ' Tinh L (Lech No) va M (Lech Co) bang Dictionary lookup (thay vi SUMPRODUCT)
            If Len(tkTrim) = 3 Or (Len(tkTrim) >= 4 And Len(tkTrim) <= 5) Or useLeftMatch Then
                ' PREFIX match: TK 3 ky tu, 4-5 ky tu, hoac TK tong hop
                pfx = CStr(Len(tkTrim)) & "|" & tkTrim
                If dictNoPrefix.Exists(pfx) Then sumNo = CDbl(dictNoPrefix(pfx)) Else sumNo = 0#
                If dictCoPrefix.Exists(pfx) Then sumCo = CDbl(dictCoPrefix(pfx)) Else sumCo = 0#
            Else
                ' EXACT match: TK chi tiet >= 6 ky tu va ton tai trong SNKC
                If dictNoExact.Exists(tkTrim) Then sumNo = CDbl(dictNoExact(tkTrim)) Else sumNo = 0#
                If dictCoExact.Exists(tkTrim) Then sumCo = CDbl(dictCoExact(tkTrim)) Else sumCo = 0#
            End If
            wsTB.Cells(r, 12).Value = sumNo - NzVal(wsTB.Cells(r, 7).Value)
            wsTB.Cells(r, 13).Value = sumCo - NzVal(wsTB.Cells(r, 8).Value)
        End If
NextTBCalc:
    Next r
    ' ========== TO MAU BANG UNION (1 lan thay vi tung dong) ==========
    ' Clear mau cu truoc
    wsTB.Range("A4:M" & lastRowTB).Interior.Pattern = xlNone
    wsTB.Range("A4:M" & lastRowTB).Font.Bold = False
    For r = 4 To lastRowTB
        If IsError(wsTB.Cells(r, 3).Value) Then GoTo NextTBColor
        If wsTB.Cells(r, 3).Value <> "" Then
            Dim tkLen As Long
            If IsNumeric(wsTB.Cells(r, 2).Value) Then tkLen = CLng(wsTB.Cells(r, 2).Value) Else tkLen = 0
            If wsTB.Cells(r, 12).Value <> 0 Or wsTB.Cells(r, 13).Value <> 0 Then
                ' Vang: dong co lech
                If rngYellow Is Nothing Then
                    Set rngYellow = wsTB.Range("A" & r & ":M" & r)
                Else
                    Set rngYellow = Union(rngYellow, wsTB.Range("A" & r & ":M" & r))
                End If
            ElseIf tkLen = 3 Then
                ' Xanh + Bold: TK cap 3
                If rngGreen Is Nothing Then
                    Set rngGreen = wsTB.Range("B" & r & ":J" & r)
                Else
                    Set rngGreen = Union(rngGreen, wsTB.Range("B" & r & ":J" & r))
                End If
                If rngBold Is Nothing Then
                    Set rngBold = wsTB.Range("A" & r & ":M" & r)
                Else
                    Set rngBold = Union(rngBold, wsTB.Range("A" & r & ":M" & r))
                End If
            End If
        End If
NextTBColor:
    Next r
    ' Ap dung mau 1 lan
    If Not rngYellow Is Nothing Then rngYellow.Interior.Color = RGB(255, 255, 150)
    If Not rngGreen Is Nothing Then rngGreen.Interior.Color = RGB(146, 208, 80)
    If Not rngBold Is Nothing Then rngBold.Font.Bold = True
    Application.Calculation = oldCalc
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
            If IsError(arr(r, c)) Then arr(r, c) = 0: GoTo NextCell
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
        If lastData >= 1 Then
            Dim arrDataL As Variant
            arrDataL = wsData.Range("L1:L" & lastData).Value
            Dim tmpTK As String
            For r = 1 To UBound(arrDataL, 1)
                tmpTK = NormalizeAccount(arrDataL(r, 1))
                If Len(tmpTK) >= 4 Then dictOpp4(Left$(tmpTK, 4)) = True
            Next r
        End If
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
        Dim tkTB As String
        lastTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
        If lastTB >= 4 Then
            Dim arrTBData As Variant
            arrTBData = wsTB.Range("C4:F" & lastTB).Value  ' C=TK, D=skip, E=DuNo, F=DuCo
            For r = 1 To UBound(arrTBData, 1)
                tkTB = NormalizeAccount(arrTBData(r, 1))
                If tkTB <> "" And Left$(tkTB, lenMain) = tkRoot Then
                    If Len(tkTB) = lenMain Then
                        duNoExact = duNoExact + NzVal(arrTBData(r, 3))
                        duCoExact = duCoExact + NzVal(arrTBData(r, 4))
                        hasExact = True
                    Else
                        duNoLeft = duNoLeft + NzVal(arrTBData(r, 3))
                        duCoLeft = duCoLeft + NzVal(arrTBData(r, 4))
                    End If
                End If
            Next r
        End If
        If hasExact Then
            duNoDK = duNoExact
            duCoDK = duCoExact
        Else
            duNoDK = duNoLeft
            duCoDK = duCoLeft
        End If
    End If
    ' Thu thap phat sinh tu NKC (array-based)
    Set dictDebit = CreateObject("Scripting.Dictionary")
    Set dictCredit = CreateObject("Scripting.Dictionary")
    Set dictDebitFull = CreateObject("Scripting.Dictionary")
    Set dictCreditFull = CreateObject("Scripting.Dictionary")
    lastNKC = wsNKC.Cells(wsNKC.Rows.Count, "A").End(xlUp).Row
    If lastNKC >= 3 Then
        Dim arrNKCTH As Variant
        arrNKCTH = wsNKC.Range("D3:H" & lastNKC).Value  ' D=TK No, E=TK Co, F=So tien, G=skip, H=Thang
        For r = 1 To UBound(arrNKCTH, 1)
            If Not hasMonthFilter Or NzVal(arrNKCTH(r, 5)) = monthFilter Then
                tkNoFull = NormalizeAccount(arrNKCTH(r, 1))
                tkCoFull = NormalizeAccount(arrNKCTH(r, 2))
                soTien = NzVal(arrNKCTH(r, 3))
                If tkNoFull <> "" And Left$(tkNoFull, lenMain) = tkRoot Then
                    If oppLenSetting > 0 And oppLenSetting >= 4 Then
                        oppKey = tkCoFull
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
    End If

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
    If IsError(v) Or IsEmpty(v) Then NormalizeAccount = "": Exit Function
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
    Dim sv As Variant
    For i = 0 To n - 1
        sv = dictSum(keys(i))
        If IsNumeric(sv) Then vals(i) = CDbl(sv) Else vals(i) = 0#
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
    ' Giam tru doanh thu: 521 No / 111, 112, 131 Co
    If noPrefix = "521" And (coPrefix = "131" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ket chuyen giam tru doanh thu: 511 No / 521 Co
    If noPrefix = "511" And coPrefix = "521" Then
        IsValidAccountPair = True: Exit Function
    End If

    ' Ket chuyen chenh lech ty gia (413) sang doanh thu tai chinh (515)
    If noPrefix = "413" And coPrefix = "515" Then
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
       (coPrefix = "334" Or coPrefix = "338") Then
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
    ' If tkNo = tkCo Then
    '     IsValidAccountPair = True: Exit Function
    ' End If
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
                    "=SUMIFS(" & wsNKC.Name & "!R3C6:R" & lastRowNKC & "C6," & wsNKC.Name & "!R3C9:R" & lastRowNKC & "C9,RC3)-RC7"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C4:R" & lastRowNKC & "C4,RC2)=RC3)*" & wsNKC.Name & "!R3C6:R" & lastRowNKC & "C6)-RC7"
            ElseIf useLeftMatch Then
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(" & wsNKC.Name & "!R3C4:R" & lastRowNKC & "C4,RC2)=RC3)*" & wsNKC.Name & "!R3C6:R" & lastRowNKC & "C6)-RC7"
            Else
                wsTB.Cells(r, 12).FormulaR1C1 = _
                    "=SUMIFS(" & wsNKC.Name & "!R3C6:R" & lastRowNKC & "C6," & wsNKC.Name & "!R3C4:R" & lastRowNKC & "C4,RC3)-RC7"
            End If
            ' Cong thuc cho cot M (Lech Co) - dung RC2 thay vi LEN(RC3) de tranh bi tinh lai khi edit
            If Len(tkTrim) = 3 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C6:R" & lastRowNKC & "C6,NKC!R3C10:R" & lastRowNKC & "C10,RC3)-RC8"
            ElseIf Len(tkTrim) >= 4 And Len(tkTrim) <= 5 Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C5:R" & lastRowNKC & "C5,RC2)=RC3)*NKC!R3C6:R" & lastRowNKC & "C6)-RC8"
            ElseIf useLeftMatch Then
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMPRODUCT((LEFT(NKC!R3C5:R" & lastRowNKC & "C5,RC2)=RC3)*NKC!R3C6:R" & lastRowNKC & "C6)-RC8"
            Else
                wsTB.Cells(r, 13).FormulaR1C1 = _
                    "=SUMIFS(NKC!R3C6:R" & lastRowNKC & "C6,NKC!R3C5:R" & lastRowNKC & "C5,RC3)-RC8"
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
    GetMonthValue = GetMonthFromAnyDate(vDate, Empty)
End Function

Private Function GetMonthFromAnyDate(ByVal vPrimary As Variant, ByVal vFallback As Variant) As Variant
    Dim d As Variant
    d = TryParseDate(vPrimary)
    If IsEmpty(d) Then d = TryParseDate(vFallback)
    If IsEmpty(d) Then
        GetMonthFromAnyDate = ""
    Else
        GetMonthFromAnyDate = Month(d)
    End If
End Function

Private Function TryParseDate(ByVal v As Variant) As Variant
    On Error GoTo Failed
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        TryParseDate = Empty
        Exit Function
    End If
    If IsDate(v) Then
        TryParseDate = CDate(v)
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        TryParseDate = Empty
        Exit Function
    End If
    ' numeric text or number -> Excel serial date or yyyymmdd
    If IsNumeric(s) Then
        Dim n As Double
        n = CDbl(s)
        If n >= 10000101# And n <= 99991231# Then
            Dim y As Long, m As Long, d As Long
            y = CLng(n \ 10000)
            m = CLng((n \ 100) Mod 100)
            d = CLng(n Mod 100)
            If y >= 1900 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                TryParseDate = DateSerial(y, m, d)
                Exit Function
            End If
        End If
        ' Excel serial date (supports numeric text like "45658")
        If n > 0 And n < 60000 Then
            TryParseDate = DateSerial(1899, 12, 30) + CLng(n)
            Exit Function
        End If
    End If
    ' normalize common separators
    s = Replace$(s, ".", "/")
    s = Replace$(s, "-", "/")
    If IsDate(s) Then
        TryParseDate = CDate(s)
        Exit Function
    End If
Failed:
    TryParseDate = Empty
End Function

Private Function NormalizeDG(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Then
        NormalizeDG = ""
        Exit Function
    End If
    t = LCase$(Trim$(CStr(s)))
    t = Replace$(t, Chr$(160), " ")
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    NormalizeDG = t
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
        If IsError(pairs(i, 3)) Then tkFull = "" Else tkFull = CStr(pairs(i, 3))
        If IsError(pairs(i, 2)) Or Not IsNumeric(pairs(i, 2)) Then amount = 0# Else amount = CDbl(pairs(i, 2))

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
    Dim dv As Variant
    For i = 0 To n - 1
        dv = dict(keys(i))
        If IsNumeric(dv) Then vals(i) = CDbl(dv) Else vals(i) = 0#
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

Private Function FindSourceSheetSmart(ByVal wb As Workbook, ByVal prefer As Worksheet) As Worksheet
    Dim ws As Worksheet
    If Not prefer Is Nothing Then
        If NameLooksLikeSource(prefer) Then
            Set FindSourceSheetSmart = prefer
            Exit Function
        End If
    End If
    For Each ws In wb.Worksheets
        If Not ws Is Nothing Then
            If LCase$(ws.Name) <> "nkc" And LCase$(ws.Name) <> "tb" And LCase$(ws.Name) <> "pv" And _
               LCase$(ws.Name) <> "pvct" And LCase$(ws.Name) <> "th" Then
                If NameLooksLikeSource(ws) Then
                    Set FindSourceSheetSmart = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
End Function

Private Function NameLooksLikeSource(ByVal ws As Worksheet) As Boolean
    Dim s As String
    s = NormalizeHeaderLocal(ws.Name)
    If s = "" Then Exit Function
    ' Ten sheet nguon thuong chua "so nhat ky chung" / "nhat ky chung"
    If InStr(s, "so nhat ky chung") > 0 Then
        NameLooksLikeSource = True
        Exit Function
    End If
    If InStr(s, "nhat ky chung") > 0 Then
        NameLooksLikeSource = True
        Exit Function
    End If
    If InStr(s, "so nhat ky") > 0 Then
        NameLooksLikeSource = True
        Exit Function
    End If
    If InStr(s, "nhat ky") > 0 Then
        NameLooksLikeSource = True
    End If
End Function

Private Function NormalizeHeaderLocal(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then NormalizeHeaderLocal = "": Exit Function
    s = LCase$(Trim$(CStr(v)))
    s = Replace$(s, Chr$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    NormalizeHeaderLocal = RemoveAccentsAsciiLocal(s)
End Function

Private Function RemoveAccentsAsciiLocal(text As String) As String
    Dim fromCodes As Variant, toChars As Variant
    Dim i As Long
    fromCodes = Array( _
        225, 224, 7843, 227, 7841, 259, 7855, 7857, 7859, 7861, 7863, _
        226, 7845, 7847, 7849, 7851, 7853, 233, 232, 7867, 7869, 7865, _
        234, 7871, 7873, 7875, 7877, 7879, 237, 236, 7881, 297, 7883, _
        243, 242, 7887, 245, 7885, 244, 7889, 7891, 7893, 7895, 7897, _
        417, 7899, 7901, 7903, 7905, 7907, 250, 249, 7911, 361, 7909, _
        432, 7913, 7915, 7917, 7919, 7921, 253, 7923, 7927, 7929, 7925, _
        273, _
        193, 192, 7842, 195, 7840, 258, 7854, 7856, 7858, 7860, 7862, _
        194, 7844, 7846, 7848, 7850, 7852, 201, 200, 7866, 7868, 7864, _
        202, 7870, 7872, 7874, 7876, 7878, 205, 204, 7880, 296, 7882, _
        211, 210, 7886, 213, 7884, 212, 7888, 7890, 7892, 7894, 7896, _
        416, 7898, 7900, 7902, 7904, 7906, 218, 217, 7910, 360, 7908, _
        431, 7912, 7914, 7916, 7918, 7920, 221, 7922, 7926, 7928, 7924, _
        272)
    toChars = Array( _
        "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", _
        "a", "a", "a", "a", "a", "a", "e", "e", "e", "e", "e", _
        "e", "e", "e", "e", "e", "e", "i", "i", "i", "i", "i", _
        "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", _
        "o", "o", "o", "o", "o", "o", "u", "u", "u", "u", "u", _
        "u", "u", "u", "u", "u", "u", "y", "y", "y", "y", "y", _
        "d", _
        "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", _
        "A", "A", "A", "A", "A", "A", "E", "E", "E", "E", "E", _
        "E", "E", "E", "E", "E", "E", "I", "I", "I", "I", "I", _
        "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", _
        "O", "O", "O", "O", "O", "O", "U", "U", "U", "U", "U", _
        "U", "U", "U", "U", "U", "U", "Y", "Y", "Y", "Y", "Y", _
        "D")
    For i = 0 To UBound(fromCodes)
        text = Replace$(text, ChrW(fromCodes(i)), toChars(i))
    Next i
    RemoveAccentsAsciiLocal = text
End Function

Private Function SheetHasDataRows(ByVal ws As Worksheet, ByVal startRow As Long) As Boolean
    Dim lastRow As Long, lastCol As Long
    On Error Resume Next
    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    On Error GoTo 0
    If lastRow < startRow Then Exit Function
    If lastCol <= 0 Then Exit Function
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol))) = 0 Then Exit Function
    SheetHasDataRows = True
End Function

Private Sub ApplyWorkbookFont(ByVal wb As Workbook, ByVal fontName As String)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        ws.Cells.Font.Name = fontName
    Next ws
    On Error GoTo 0
End Sub

Private Sub ReorderNKCColumnsIfOld(ws As Worksheet)
    On Error Resume Next
    Dim h1 As String, h2 As String
    h1 = LCase$(Trim$(CStr(ws.Cells(2, 1).Value)))
    h2 = LCase$(Trim$(CStr(ws.Cells(2, 2).Value)))
    On Error GoTo 0
    ' Old layout: A=Ngay hoach toan, B=Ngay chung tu
    If InStr(h1, "ng") = 1 And InStr(h1, "hoach") > 0 And InStr(h2, "chung") > 0 Then
        Dim lastRow As Long, lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        If lastRow < 2 Or lastCol < 10 Then Exit Sub
        Dim src As Variant, dst As Variant
        Dim r As Long, c As Long, idx As Long
        src = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
        ReDim dst(1 To UBound(src, 1), 1 To lastCol)
        ' New order base: 2,4,5,8,9,10,1,3,6,7 then keep 11..last
        Dim map As Variant
        map = Array(2, 4, 5, 8, 9, 10, 1, 3, 6, 7)
        For r = 1 To UBound(src, 1)
            For idx = 0 To UBound(map)
                dst(r, idx + 1) = src(r, map(idx))
            Next idx
            If lastCol > 10 Then
                For c = 11 To lastCol
                    dst(r, c) = src(r, c)
                Next c
            End If
        Next r
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value = dst
    End If
End Sub
