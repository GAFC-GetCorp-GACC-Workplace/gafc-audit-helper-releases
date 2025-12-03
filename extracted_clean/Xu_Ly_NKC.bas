Option Explicit
Public Sub Test_Xu_ly_NKC()
    Xu_ly_NKC1111 Nothing
End Sub
Public Sub Chinh_Format_NKC_va_Pivot(control As IRibbonControl)
    Dim errs As String
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
    Dim wb As Workbook
    Dim wsNguon As Worksheet
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsNguon = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0
    If wsNguon Is Nothing Then Set wsNguon = ActiveSheet
    wsNguon.Activate
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
    Dim wsNguon As Worksheet, wsKetQua As Worksheet
    Dim wb As Workbook
    Dim dictGroup As Object
    Dim lastRow As Long, i As Long
    Dim arrData As Variant
    Dim key As Variant, r As Variant
    Dim pivotErr As String, thMsg As String
    Set wb = ActiveWorkbook
    ' Uu tien sheet "So Nhat Ky Chung" neu co, bat ke dang dung sheet nao
    On Error Resume Next
    Set wsNguon = wb.Sheets("So Nhat Ky Chung")
    On Error GoTo 0
    If wsNguon Is Nothing Then Set wsNguon = ActiveSheet
    If wsNguon Is Nothing Then
        MsgBox "Kh" & ChrW(244) & "ng x" & ChrW(225) & "c " & ChrW(273) & ChrW(7883) & "nh " & ChrW(273) & ChrW(432) & ChrW(7907) & "c sheet ngu" & ChrW(7891) & "n!", vbExclamation
        Exit Sub
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
        .Cells(2, 11).Value = "C" & ChrW(7847) & "n review"
        .Range("A2:K2").Font.Bold = True
        .Range("A2:K2").Interior.Color = RGB(220, 230, 241)
        .Range("A2:K2").AutoFilter
        .Columns("A:K").AutoFit
    End With
    Set dictGroup = CreateObject("Scripting.Dictionary")
    ' Nhom du lieu theo MaCT|Ngay
    For i = 1 To UBound(arrData, 1)
        If Trim(arrData(i, 1)) <> "" Then
            key = Trim(arrData(i, 1)) & "|" & Trim(arrData(i, 2))
            If Not dictGroup.Exists(key) Then dictGroup.Add key, New Collection
            dictGroup(key).Add i
        End If
    Next i
    ' ========== BUOC 1: XAC DINH NHOM "BAN" ==========
    ' Nhom "ban" = co it nhat 1 dong co CA No va Co
    Dim dictDirty As Object
    Set dictDirty = CreateObject("Scripting.Dictionary")
    For Each key In dictGroup.keys
        Dim isDirty As Boolean
        isDirty = False
        For Each r In dictGroup(key)
            If val(arrData(r, 5)) <> 0 And val(arrData(r, 6)) <> 0 Then
                isDirty = True
                Exit For
            End If
        Next r
        dictDirty.Add key, isDirty
    Next key
    ' ========== XU LY VA THU THAP OUTPUT ==========
    Dim outputArr() As Variant
    ReDim outputArr(1 To (lastRow - 1) * 10, 1 To 11)
    Dim dongOut As Long
    dongOut = 1
    For Each key In dictGroup.keys
        Dim dsNoEntries As Collection, dsCoEntries As Collection
        Set dsNoEntries = New Collection
        Set dsCoEntries = New Collection
        For Each r In dictGroup(key)
            Dim tienNoGoc As Double, tienCoGoc As Double
            tienNoGoc = val(arrData(r, 5))
            tienCoGoc = val(arrData(r, 6))
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
        Dim tkNo As String, tkCo As String
        Dim needReview As String
        ' Lay trang thai "ban" cua nhom
        needReview = ""
        If dictDirty(key) Then needReview = "X"
        ' ========== PASS 1: Ghep theo QUY TAC KE TOAN ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass1
            tkNo = CStr(arrData(rNo, 4))
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoPass1
                tkCo = CStr(arrData(rCo, 4))
                If Abs(tienNo - tienCo) < 0.01 And IsValidAccountPair(tkNo, tkCo) Then
                    If dongOut > UBound(outputArr, 1) Then
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To 11)
                    End If
                    ' Format mau: NgayHT, NgayCT, Thang, SoHD, DienGiai, No, Co, NoTK, CoTK, SoTien, CanReview
                    outputArr(dongOut, 1) = arrData(rNo, 2)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    ' Thang lay theo Ngay hach toan (cot A sheet NKC = col 2 nguon)
                    outputArr(dongOut, 3) = GetMonthValue(arrData(rNo, 2))
                    outputArr(dongOut, 4) = arrData(rNo, 1)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrData(rNo, 3)  ' Dien giai
                    outputArr(dongOut, 6) = Left(CStr(arrData(rNo, 4)), 3)  ' No (3 ky tu dau TK No)
                    outputArr(dongOut, 7) = Left(CStr(arrData(rCo, 4)), 3)  ' Co (3 ky tu dau TK Co)
                    outputArr(dongOut, 8) = arrData(rNo, 4)  ' No TK (full)
                    outputArr(dongOut, 9) = arrData(rCo, 4)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienNo          ' So tien
                    outputArr(dongOut, 11) = needReview      ' CanReview
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
                        ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To 11)
                    End If
                    outputArr(dongOut, 1) = arrData(rNo, 2)  ' Ngay hach toan
                    outputArr(dongOut, 2) = ""               ' Ngay chung tu
                    outputArr(dongOut, 3) = GetMonthValue(arrData(rNo, 2))
                    outputArr(dongOut, 4) = arrData(rNo, 1)  ' So hoa don (MaCT)
                    outputArr(dongOut, 5) = arrData(rNo, 3)  ' Dien giai
                    outputArr(dongOut, 6) = Left(CStr(arrData(rNo, 4)), 3)  ' No (3 ky tu dau)
                    outputArr(dongOut, 7) = Left(CStr(arrData(rCo, 4)), 3)  ' Co (3 ky tu dau)
                    outputArr(dongOut, 8) = arrData(rNo, 4)  ' No TK (full)
                    outputArr(dongOut, 9) = arrData(rCo, 4)  ' Co TK (full)
                    outputArr(dongOut, 10) = tienNo          ' So tien
                    outputArr(dongOut, 11) = needReview      ' CanReview
                    usedNo(idxNo) = usedNo(idxNo) + tienNo
                    usedCo(idxCo) = usedCo(idxCo) + tienNo
                    dongOut = dongOut + 1
                    Exit For
                End If
NextCoPass2:
            Next idxCo
NextNoPass2:
        Next idxNo
        ' ========== PASS 3: Phan bo phan con lai ==========
        For idxNo = 1 To dsNoEntries.Count
            entryNo = dsNoEntries(idxNo)
            rNo = entryNo(0)
            tienNoEntry = entryNo(1)
            tienNo = tienNoEntry - usedNo(idxNo)
            If Abs(tienNo) < 0.01 Then GoTo NextNoPass3
            For idxCo = 1 To dsCoEntries.Count
                entryCo = dsCoEntries(idxCo)
                rCo = entryCo(0)
                tienCoEntry = entryCo(1)
                tienCo = tienCoEntry - usedCo(idxCo)
                If Abs(tienCo) < 0.01 Then GoTo NextCoPass3
                tienPhanBo = Application.Min(Abs(tienNo), Abs(tienCo)) * Sgn(tienNo)
                If dongOut > UBound(outputArr, 1) Then
                    ReDim Preserve outputArr(1 To UBound(outputArr, 1) * 2, 1 To 11)
                End If
                outputArr(dongOut, 1) = arrData(rNo, 2)  ' Ngay hach toan
                outputArr(dongOut, 2) = ""               ' Ngay chung tu
                outputArr(dongOut, 3) = GetMonthValue(arrData(rNo, 2))
                outputArr(dongOut, 4) = arrData(rNo, 1)  ' So hoa don (MaCT)
                outputArr(dongOut, 5) = arrData(rNo, 3)  ' Dien giai
                outputArr(dongOut, 6) = Left(CStr(arrData(rNo, 4)), 3)  ' No (3 ky tu dau)
                outputArr(dongOut, 7) = Left(CStr(arrData(rCo, 4)), 3)  ' Co (3 ky tu dau)
                outputArr(dongOut, 8) = arrData(rNo, 4)  ' No TK (full)
                outputArr(dongOut, 9) = arrData(rCo, 4)  ' Co TK (full)
                outputArr(dongOut, 10) = tienPhanBo      ' So tien
                outputArr(dongOut, 11) = needReview      ' CanReview
                usedNo(idxNo) = usedNo(idxNo) + tienPhanBo
                usedCo(idxCo) = usedCo(idxCo) + tienPhanBo
                tienNo = tienNo - tienPhanBo
                dongOut = dongOut + 1
                If Abs(tienNo) < 0.01 Then Exit For
NextCoPass3:
            Next idxCo
NextNoPass3:
        Next idxNo
NextGroup:
    Next key
    ' ========== GHI OUTPUT ==========
    If dongOut > 1 Then
        Dim finalOut() As Variant
        ReDim finalOut(1 To dongOut - 1, 1 To 11)
        Dim j As Long
        For i = 1 To dongOut - 1
            For j = 1 To 11
                finalOut(i, j) = outputArr(i, j)
            Next j
        Next i
        wsKetQua.Range("A3").Resize(dongOut - 1, 11).Value = finalOut
        ' Tinh thang tu cot Ngay hach toan (col A) bang MONTH, sau do fix value
        Dim dataLast As Long
        dataLast = dongOut + 1 ' hang cuoi co du lieu (bat dau tu row 3)
        With wsKetQua
            .Range("C3:C" & dataLast).FormulaR1C1 = "=MONTH(RC[-2])"
            .Range("C3:C" & dataLast).Value = .Range("C3:C" & dataLast).Value
        End With
        ' ========== TO VANG CAC DONG CAN REVIEW ==========
        Dim rng As Range
        For i = 3 To dongOut + 1
            If wsKetQua.Cells(i, 11).Value = "X" Then
                Set rng = wsKetQua.Range(wsKetQua.Cells(i, 1), wsKetQua.Cells(i, 11))
                rng.Interior.Color = RGB(255, 255, 150) ' Vang
            End If
        Next i
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
    countReview = Application.WorksheetFunction.CountIf(wsKetQua.Columns(11), "X")
    ' Dem so nhom ban
    Dim countDirty As Long
    countDirty = 0
    For Each key In dictDirty.keys
        If dictDirty(key) Then countDirty = countDirty + 1
    Next key
    ' Bat lai cac tinh nang truoc khi tinh TB
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ' Hien thi ket qua NKC truoc
    MsgBox "X" & ChrW(7916) & " L" & ChrW(221) & " NKC HO" & ChrW(192) & "N TH" & ChrW(192) & "NH!" & vbCrLf & vbCrLf & _
           "T" & ChrW(7893) & "ng s" & ChrW(7889) & " b" & ChrW(250) & "t to" & ChrW(225) & "n output: " & (dongOut - 1) & vbCrLf & _
           "S" & ChrW(7889) & " nh" & ChrW(243) & "m B" & ChrW(218) & "T TO" & ChrW(193) & "N B" & ChrW(7848) & "N: " & countDirty & " ch" & ChrW(7913) & "ng t" & ChrW(7915) & vbCrLf & _
           "S" & ChrW(7889) & " d" & ChrW(242) & "ng C" & ChrW(7846) & "N REVIEW (t" & ChrW(244) & " v" & ChrW(224) & "ng): " & countReview, vbInformation
    ' Sau do moi tinh TB (neu co)
    Dim tbMsg As String
    If WorksheetExists("TB", wb) Then
        tbMsg = Auto_Tinh_TB(wsKetQua)
        If tbMsg <> "" Then
            MsgBox "T" & ChrW(205) & "NH TO" & ChrW(193) & "N TB TH" & ChrW(192) & "NH C" & ChrW(212) & "NG!" & tbMsg, vbInformation
        End If
    Else
        MsgBox "C" & ChrW(7842) & "NH B" & ChrW(193) & "O: Kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y sheet TB!" & vbCrLf & vbCrLf & _
               "Vui l" & ChrW(242) & "ng t" & ChrW(7841) & "o m" & ChrW(7851) & "u c" & ChrW(243) & " sheet TB " & ChrW(273) & ChrW(7875) & " t" & ChrW(7921) & " " & ChrW(273) & ChrW(7897) & "ng t" & ChrW(237) & "nh to" & ChrW(225) & "n.", vbExclamation
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
    If thMsg <> "" Then MsgBox thMsg, vbExclamation
    ' Bat auto refresh TH sau khi da co NKC/TB
    On Error Resume Next
    Application.Run "Enable_TH_AutoRefresh"
    On Error GoTo 0
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
    On Error GoTo ErrorHandler
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
    Application.Calculation = xlCalculationAutomatic
    wsTB.Calculate
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
    wsTB.Range("E1:M" & lastRowTB).NumberFormat = "#,##0"
    Auto_Tinh_TB = ""
    Exit Function
ErrorHandler:
    MsgBox "L" & ChrW(7894) & "I: Kh" & ChrW(244) & "ng th" & ChrW(7875) & " t" & ChrW(237) & "nh to" & ChrW(225) & "n TB!" & vbCrLf & vbCrLf & _
           "Chi ti" & ChrW(7871) & "t: " & Err.Description, vbCritical
    Auto_Tinh_TB = ""
End Function
' Wrapper cho Ribbon button - Cap nhat dropdown TH
Public Sub Update_TH_Dropdown_Button(control As IRibbonControl)
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
    End If
    If wsTH Is Nothing Then
        Auto_Tinh_TH = "Khong the tao sheet TH."
        Exit Function
    End If
    tkRaw = NormalizeAccount(wsTH.Range("C4").Value)
    If tkRaw = "" Then
        Auto_Tinh_TH = "" ' silent if chua nhap TK
        Exit Function
    End If
    ' So dong doi ung hien thi (giu 6 dong de khong de len dong SPS/SDCK)
    slotCount = 6
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
    ' Clear vung du lieu cu (toi da 12 dong doi ung -> hang 6..17)
    wsTH.Range("B6:E17").ClearContents
    wsTH.Range("I7:K18").ClearContents
    wsTH.Range("I19:K19").ClearContents
    ' Ghi phat sinh No (TK goc o ben No -> doi ung o col B/C)
    If dictDebit.Count > 0 Then
        pairs = SortDictByAbsWithFull(dictDebit, dictDebitFull)
        Dim maxDebit As Long
        maxDebit = slotCount
        If UBound(pairs, 1) < maxDebit Then maxDebit = UBound(pairs, 1)
        For i = 1 To maxDebit
            Dim dispOppN As String
            If oppLenSetting > 0 And oppLenSetting <= 3 Then
                dispOppN = pairs(i, 1) ' key rut gon
            Else
                dispOppN = pairs(i, 3) ' full
            End If
            wsTH.Cells(5 + i, 2).Value = IIf(dispOppN <> "", "<" & dispOppN & ">", "")
            wsTH.Cells(5 + i, 3).Value = pairs(i, 2) ' so tien
        Next i
    End If
    ' Ghi phat sinh Co (TK goc o ben Co -> doi ung o col D/E)
    If dictCredit.Count > 0 Then
        pairs = SortDictByAbsWithFull(dictCredit, dictCreditFull)
        Dim maxCredit As Long
        maxCredit = slotCount
        If UBound(pairs, 1) < maxCredit Then maxCredit = UBound(pairs, 1)
        For i = 1 To maxCredit
            Dim dispOppC As String
            If oppLenSetting > 0 And oppLenSetting <= 3 Then
                dispOppC = pairs(i, 1)
            Else
                dispOppC = pairs(i, 3)
            End If
            wsTH.Cells(5 + i, 4).Value = pairs(i, 2) ' so tien
            wsTH.Cells(5 + i, 5).Value = IIf(dispOppC <> "", "<" & dispOppC & ">", "")
        Next i
    End If
    ' Ghi SPS va SDCK theo T-account
    wsTH.Range("C12").Value = totalDebitPS
    wsTH.Range("D12").Value = totalCreditPS
    sdBalance = (duNoDK - duCoDK) + (totalDebitPS - totalCreditPS)
    wsTH.Range("C13").Value = Application.Max(sdBalance, 0)
    wsTH.Range("D13").Value = Application.Max(-sdBalance, 0)
    Auto_Tinh_TH = ""
    Exit Function
ErrHandler:
    Auto_Tinh_TH = "Khong the cap nhat sheet TH. Chi tiet: " & Err.Description
End Function
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
Function IsValidAccountPair(tkNo As String, tkCo As String) As Boolean
    Dim noPrefix As String, coPrefix As String
    noPrefix = Left(tkNo, 3)
    coPrefix = Left(tkCo, 3)
    IsValidAccountPair = False
    ' === QUY TAC KET CHUYEN 911 ===
    If noPrefix = "911" And (Left(tkCo, 1) = "6" Or Left(tkCo, 1) = "8") Then
        IsValidAccountPair = True: Exit Function
    End If
    If (Left(tkNo, 1) = "5" Or Left(tkNo, 1) = "7") And coPrefix = "911" Then
        IsValidAccountPair = True: Exit Function
    End If
    If (noPrefix = "911" And coPrefix = "421") Or (noPrefix = "421" And coPrefix = "911") Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC KET CHUYEN CHI PHI SAN XUAT 154 ===
    If noPrefix = "154" And (coPrefix = "621" Or coPrefix = "622" Or coPrefix = "627") Then
        IsValidAccountPair = True: Exit Function
    End If
    ' 155 No / 154 Co (nhap thanh pham tu SPDD)
    If noPrefix = "155" And coPrefix = "154" Then
        IsValidAccountPair = True: Exit Function
    End If
    ' 632 No / 154, 155, 156 Co (gia von)
    If noPrefix = "632" And (coPrefix = "154" Or coPrefix = "155" Or coPrefix = "156") Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC MUA HANG ===
    If (noPrefix = "152" Or noPrefix = "153" Or noPrefix = "156") And _
       (coPrefix = "331" Or coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If
    If noPrefix = "133" And coPrefix = "331" Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC BAN HANG ===
    If (noPrefix = "131" Or noPrefix = "111" Or noPrefix = "112") And _
       (coPrefix = "511" Or coPrefix = "333") Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC THANH TOAN ===
    If noPrefix = "331" And (coPrefix = "111" Or coPrefix = "112") Then
        IsValidAccountPair = True: Exit Function
    End If
    If (noPrefix = "111" Or noPrefix = "112") And coPrefix = "131" Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC LUONG & BAO HIEM ===
    If (noPrefix = "622" Or noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And _
       coPrefix = "334" Then
        IsValidAccountPair = True: Exit Function
    End If
    If noPrefix = "334" And (coPrefix = "111" Or coPrefix = "112" Or coPrefix = "338") Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === QUY TAC KHAU HAO ===
    If (noPrefix = "627" Or noPrefix = "641" Or noPrefix = "642") And coPrefix = "214" Then
        IsValidAccountPair = True: Exit Function
    End If
    ' === CUNG TAI KHOAN (but toan noi bo) ===
    If tkNo = tkCo Then
        IsValidAccountPair = True: Exit Function
    End If
End Function
' ===================================================================
' Ham Tinh Toan TB (Trial Balance)
' ===================================================================
Public Sub Tinh_Toan_TB(control As IRibbonControl)
    Dim wsNKC As Worksheet, wsTB As Worksheet, wsSource As Worksheet
    Dim wb As Workbook
    Dim lastRowNKC As Long, lastRowTB As Long, lastRowSource As Long
    Dim r As Long, i As Long
    Dim tkFull As String, tkTrim As String
    Dim dictSourceTK As Object
    Dim tkSource As String
    Dim useLeftMatch As Boolean
    Set wb = ActiveWorkbook
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
    Application.Calculation = xlCalculationAutomatic
    wsTB.Calculate
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
    wsTB.Range("E1:M" & lastRowTB).NumberFormat = "#,##0"
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
