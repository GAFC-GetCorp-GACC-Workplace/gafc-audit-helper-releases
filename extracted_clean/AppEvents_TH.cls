VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents_TH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo ExitHandler
    If Sh Is Nothing Then Exit Sub
    If Sh.Name <> "TH" Then Exit Sub
    ' Cap nhat dropdown C4 khi B4/D3 thay doi
    If Not Intersect(target, Sh.Range("B4,D3")) Is Nothing Then
        Application.EnableEvents = False
        Update_TH_Dropdown_Internal Sh
        Application.EnableEvents = True
    End If
    ' Refresh TH khi C4, D2, D3, D1, D4 thay doi
    If Intersect(target, Sh.Range("C4,D2,D3,D1,D4")) Is Nothing Then Exit Sub
    Refresh_TH Sh.Parent
ExitHandler:
    On Error Resume Next
End Sub
Private Sub Update_TH_Dropdown_Internal(ByVal wsTH As Object)
    Dim wsTB As Worksheet
    Dim tkPrefix As String
    Dim dictTK As Object
    Dim dictLen As Object
    Dim lastRowTB As Long, r As Long
    Dim tkList As String, tkItem As String
    Dim keys As Variant, i As Long
    Dim helperRange As Range
    Dim capLen As Long
    Dim hasCap As Boolean
    Dim firstVal As String
    Dim lenList As Variant
    Dim fallbackLen As Long
    Dim needRebuild As Boolean
    On Error Resume Next
    Set wsTB = wsTH.Parent.Sheets("TB")
    If wsTB Is Nothing Then Exit Sub
    ' Lay TK prefix tu B4
    tkPrefix = Trim$(CStr(wsTH.Range("B4").Value))
    ' Loai bo apostrophe neu co
    If Left$(tkPrefix, 1) = "'" Then tkPrefix = Mid$(tkPrefix, 2)
    tkPrefix = Trim$(tkPrefix)
    If tkPrefix = "" Then
        wsTH.Range("C4").Validation.Delete
        Exit Sub
    End If
    ' Cap TK (so ky tu) tu D3 (neu co)
    capLen = 0
    If IsNumeric(wsTH.Range("D3").Value) Then
        capLen = CLng(wsTH.Range("D3").Value)
        If capLen > 0 Then hasCap = True
    End If
    ' Thu thap tat ca TK con tu TB
    Set dictTK = CreateObject("Scripting.Dictionary")
    Set dictLen = CreateObject("Scripting.Dictionary")
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "C").End(xlUp).Row
    ' Doc tat ca du lieu vao array de tranh bi anh huong boi filter
    Dim arrTB As Variant
    If lastRowTB >= 4 Then
        arrTB = wsTB.Range("C4:C" & lastRowTB).Value
        For r = 1 To UBound(arrTB, 1)
            tkItem = Trim$(CStr(arrTB(r, 1)))
            ' Loai bo apostrophe neu co
            If Left$(tkItem, 1) = "'" Then tkItem = Mid$(tkItem, 2)
            tkItem = Trim$(tkItem)
            If tkItem <> "" And Left$(tkItem, Len(tkPrefix)) = tkPrefix Then
                If Not dictLen.Exists(Len(tkItem)) Then dictLen.Add Len(tkItem), True
                If hasCap And Len(tkItem) <> capLen Then GoTo NextR
                If Not dictTK.Exists(tkItem) Then
                    dictTK.Add tkItem, True
                End If
            End If
NextR:
        Next r
    End If
    ' Neu co capLen ma khong co ket qua thi fallback ve cap 3 neu co, nguoc lai lay cap nho nhat
    If hasCap And dictTK.Count = 0 And dictLen.Count > 0 Then
        ' Tao danh sach cap
        lenList = dictLen.keys
        ' Tim cap fallback
        fallbackLen = lenList(0)
        For i = 0 To UBound(lenList)
            If lenList(i) = 3 Then
                fallbackLen = 3
                Exit For
            ElseIf lenList(i) < fallbackLen Then
                fallbackLen = lenList(i)
            End If
        Next i
        capLen = fallbackLen
        hasCap = True
        needRebuild = True
        ' Gan D3 theo cap fallback
        wsTH.Range("D3").NumberFormat = "0"
        wsTH.Range("D3").Value = capLen
    End If
    ' Tao validation cho D3 tu danh sach cap ton tai
    If dictLen.Count > 0 Then
        lenList = dictLen.keys
        ' sort tang dan
        Dim j As Long, k As Long, tmp As Variant
        For j = LBound(lenList) To UBound(lenList) - 1
            For k = j + 1 To UBound(lenList)
                If lenList(k) < lenList(j) Then
                    tmp = lenList(j): lenList(j) = lenList(k): lenList(k) = tmp
                End If
            Next k
        Next j
        tkList = lenList(0)
        For i = 1 To UBound(lenList)
            tkList = tkList & "," & lenList(i)
        Next i
        With wsTH.Range("D3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:=tkList
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Else
        wsTH.Range("D3").Validation.Delete
    End If
    ' Neu can rebuild danh sach TK theo cap fallback
    If needRebuild Then
        dictTK.RemoveAll
        For r = 1 To UBound(arrTB, 1)
            tkItem = Trim$(CStr(arrTB(r, 1)))
            If Left$(tkItem, 1) = "'" Then tkItem = Mid$(tkItem, 2)
            tkItem = Trim$(tkItem)
            If tkItem <> "" And Left$(tkItem, Len(tkPrefix)) = tkPrefix Then
                If Len(tkItem) = capLen Then
                    If Not dictTK.Exists(tkItem) Then dictTK.Add tkItem, True
                End If
            End If
        Next r
    End If
    ' Tao list cho validation
    If dictTK.Count > 0 Then
        keys = dictTK.keys
        firstVal = keys(0)
        tkList = keys(0)
        For i = 1 To UBound(keys)
            tkList = tkList & "," & keys(i)
        Next i
        ' Neu list qua dai (>255 ky tu), dung Named Range thay the
        If Len(tkList) > 255 Then
            Set helperRange = wsTH.Range("Z1").Resize(dictTK.Count, 1)
            helperRange.ClearContents
            For i = 0 To UBound(keys)
                helperRange.Cells(i + 1, 1).Value = keys(i)
            Next i
            With wsTH.Range("C4").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                     Formula1:="=" & helperRange.Address
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Else
            With wsTH.Range("C4").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                     Operator:=xlBetween, Formula1:=tkList
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End If
        ' Dat gia tri mac dinh la muc dau tien
        wsTH.Range("C4").NumberFormat = "@"
        wsTH.Range("C4").Value = "'" & firstVal
        ' Sau khi set gia tri, refresh TH de cap nhat ngay
        On Error Resume Next
        Refresh_TH wsTH.Parent
        On Error GoTo 0
    Else
        wsTH.Range("C4").Validation.Delete
        wsTH.Range("C4").ClearContents
    End If
    On Error GoTo 0
End Sub
Private Sub App_SheetActivate(ByVal Sh As Object)
    ' Tat auto-refresh khi activate de tranh lag
    ' Chi refresh khi thay doi cell (App_SheetChange)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If Sh.Name = "TH" Then
        EnsureBoolValidation Sh.Range("D1,D4")
        EnsureMonthValidation Sh.Range("D2")
    End If
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    Set App = Nothing
End Sub
Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal target As Range, Cancel As Boolean)
    HandleDoubleClickCore Sh, target, Cancel
End Sub
Public Sub Handle_TH_DoubleClick(ByVal Sh As Object, ByVal target As Range, Cancel As Boolean)
    HandleDoubleClickCore Sh, target, Cancel
End Sub
Private Sub HandleDoubleClickCore(ByVal Sh As Object, ByVal target As Range, ByRef Cancel As Boolean)
    Dim tkRoot As String, oppLabel As String, oppKey As String
    Dim isDebitSide As Boolean
    Dim monthVal As Variant, hasMonth As Boolean
    Dim oppLenSetting As Long
    On Error GoTo ExitHandler
    If Sh Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If Sh.Name <> "TH" Then Exit Sub
    ' Toggle TRUE/FALSE cho D1, D4 bang double-click
    If target.CountLarge = 1 Then
        If target.Address(False, False) = "D1" Or target.Address(False, False) = "D4" Then
            ToggleBooleanAndRefresh Sh, target
            Cancel = True
            Exit Sub
        End If
    End If
    If target.CountLarge <> 1 Then Exit Sub
    If Intersect(target, Sh.Range("C:D")) Is Nothing Then Exit Sub
    If target.Row < 5 Then Exit Sub
    If Not IsNumeric(target.Value) Then Exit Sub

    tkRoot = NormalizeAccountLocal(Sh.Range("C4").Value)
    If tkRoot = "" Then Exit Sub
    If IsNumeric(Sh.Range("D2").Value) Then
        If CLng(Sh.Range("D2").Value) > 0 Then
            hasMonth = True
            monthVal = CLng(Sh.Range("D2").Value)
        End If
    End If

    If target.Column = 3 Then
        isDebitSide = True
        oppLabel = Sh.Cells(target.Row, 2).Value
    ElseIf target.Column = 4 Then
        isDebitSide = False
        oppLabel = Sh.Cells(target.Row, 5).Value
    Else
        Exit Sub
    End If
    oppLabel = CleanOppLabel(oppLabel)
    If oppLabel = "" Then Exit Sub

    ' Dung full TK doi ung (khong rut gon)
    oppKey = oppLabel

    ' Cap (so ky tu) tu D3 neu co
    If IsNumeric(Sh.Range("D3").Value) Then
        oppLenSetting = CLng(Sh.Range("D3").Value)
        If oppLenSetting < 0 Then oppLenSetting = 0
    End If

    FilterNKCFromTH Sh.Parent, tkRoot, oppKey, isDebitSide, hasMonth, monthVal, oppLenSetting
    Cancel = True
ExitHandler:
End Sub
Private Sub ToggleBooleanAndRefresh(ws As Worksheet, cell As Range)
    On Error Resume Next
    Application.EnableEvents = False
    If CBoolSafe(cell.Value) Then
        cell.Value = False
    Else
        cell.Value = True
    End If
    Application.EnableEvents = True
    Refresh_TH ws.Parent
End Sub
Private Function CBoolSafe(v As Variant) As Boolean
    On Error Resume Next
    CBoolSafe = False
    If VarType(v) = vbBoolean Then
        CBoolSafe = CBool(v)
    ElseIf VarType(v) = vbString Then
        If UCase$(Trim$(CStr(v))) = "TRUE" Then CBoolSafe = True
    ElseIf IsNumeric(v) Then
        CBoolSafe = (CDbl(v) <> 0)
    End If
    On Error GoTo 0
End Function
Private Function CleanOppLabel(v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    s = Replace$(s, "<", "")
    s = Replace$(s, ">", "")
    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    CleanOppLabel = Trim$(s)
End Function
Private Function NormalizeAccountLocal(v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then NormalizeAccountLocal = "": Exit Function
    s = Trim$(CStr(v))
    If s = "" Then
        NormalizeAccountLocal = ""
        Exit Function
    End If
    ' Bo apostrophe neu co (giu nguyen gia tri text)
    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    If Right$(s, 2) = ".0" Then s = Left$(s, Len(s) - 2)
    NormalizeAccountLocal = s
End Function
Private Function GetPrefixLenFromDictLocal(tk As String, dict4 As Object, Optional overrideLen As Long = 0) As Long
    Dim safeLen As Long
    safeLen = Len(tk)
    If safeLen = 0 Then
        GetPrefixLenFromDictLocal = 0
        Exit Function
    End If
    If overrideLen > 0 Then
        If overrideLen > safeLen Then overrideLen = safeLen
        GetPrefixLenFromDictLocal = overrideLen
        Exit Function
    End If
    If safeLen >= 4 And Not dict4 Is Nothing Then
        If dict4.Exists(Left$(tk, 4)) Then
            GetPrefixLenFromDictLocal = 4
            Exit Function
        End If
    End If
    If safeLen < 3 Then
        GetPrefixLenFromDictLocal = safeLen
    Else
        GetPrefixLenFromDictLocal = 3
    End If
End Function
Private Function BuildOpp4Dict(wb As Workbook) As Object
    Dim wsData As Worksheet
    Dim lastData As Long, r As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsData = wb.Sheets("Data")
    On Error GoTo 0
    If wsData Is Nothing Then
        Set BuildOpp4Dict = dict
        Exit Function
    End If
    lastData = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    For r = 1 To lastData
        Dim tk As String
        tk = NormalizeAccountLocal(wsData.Cells(r, "L").Value)
        If Len(tk) >= 4 Then dict(Left$(tk, 4)) = True
    Next r
    Set BuildOpp4Dict = dict
End Function
Private Function ResolveOppKey(oppLabel As String, oppLenSetting As Long, dictOpp4 As Object, ByRef isExact As Boolean) As String
    Dim prefixLen As Long
    isExact = False
    If oppLabel = "" Then Exit Function
    If oppLenSetting >= 4 Then
        ResolveOppKey = oppLabel
        isExact = True
        Exit Function
    End If
    If oppLenSetting > 0 Then
        prefixLen = oppLenSetting
        If prefixLen > Len(oppLabel) Then prefixLen = Len(oppLabel)
        ResolveOppKey = Left$(oppLabel, prefixLen)
        Exit Function
    End If
    prefixLen = GetPrefixLenFromDictLocal(oppLabel, dictOpp4, oppLenSetting)
    If prefixLen >= Len(oppLabel) Then
        ResolveOppKey = oppLabel
        isExact = True
    Else
        ResolveOppKey = Left$(oppLabel, prefixLen)
    End If
End Function
Private Function BuildFilterArray(ws As Worksheet, fieldNum As Long, prefix As String, lastRow As Long) As Variant()
    Dim dict As Object
    Dim arrData As Variant
    Dim i As Long
    Dim tkVal As String
    Dim tkValOriginal As Variant
    Dim result() As Variant
    Dim prefixLen As Long
    Dim keyBase As String
    Dim itemVals As Variant
    Dim storeVal As Variant
    Dim storeValText As String

    Set dict = CreateObject("Scripting.Dictionary")
    prefixLen = Len(prefix)
    If prefixLen = 0 Then
        ReDim result(0 To 0)
        result(0) = Empty
        BuildFilterArray = result
        Exit Function
    End If

    If lastRow > 2 Then
        arrData = ws.Range(ws.Cells(3, fieldNum), ws.Cells(lastRow, fieldNum)).Value
        For i = 1 To UBound(arrData, 1)
            If Not IsEmpty(arrData(i, 1)) Then
                tkValOriginal = arrData(i, 1)
                tkVal = Trim$(CStr(tkValOriginal))
                If tkVal <> "" Then
                    If Left$(tkVal, 1) = "'" Then tkVal = Mid$(tkVal, 2)
                    If Len(tkVal) >= prefixLen Then
                        If Left$(tkVal, prefixLen) = prefix Then
                            keyBase = tkVal
                            If IsNumeric(tkValOriginal) Then
                                storeVal = CDbl(tkValOriginal) ' giu kieu so de AutoFilter khop so
                            Else
                                storeVal = CStr(tkValOriginal)
                            End If
                            storeValText = CStr(tkVal)
                            ' luu ca dang so va dang text de khop du lieu loai text/so
                            If Not dict.Exists(keyBase & "|num") And IsNumeric(tkValOriginal) Then dict.Add keyBase & "|num", storeVal
                            If Not dict.Exists(keyBase & "|str") Then dict.Add keyBase & "|str", storeValText
                        End If
                    End If
                End If
            End If
        Next i
    End If

    ' IMPORTANT: Base-1 array for Excel AutoFilter
    If dict.Count > 0 Then
        itemVals = dict.Items
        ReDim result(1 To dict.Count)
        For i = 0 To UBound(itemVals)
            result(i + 1) = itemVals(i)
        Next i
    Else
        ReDim result(0 To 0)
        result(0) = Empty
    End If

    BuildFilterArray = result
End Function

Private Sub FilterNKCFromTH(wb As Workbook, tkRoot As String, oppKey As String, isDebitSide As Boolean, hasMonth As Boolean, monthVal As Variant, oppLenSetting As Long)
    Dim wsNKC As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim fieldNo As Long, fieldCo As Long
    Dim rootField As Long, oppField As Long
    Dim amountField As Long
    Dim monthField As Long
    Dim oppPrefix As String
    Dim oldCalc As Long, oldEvents As Boolean, oldScreen As Boolean
    Dim appliedRoot As Boolean, appliedOpp As Boolean
    Dim visCount As Long

    On Error GoTo CleanUp

    ' Tat screen updating va events de tang toc
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsNKC = wb.Sheets("NKC")
    If wsNKC Is Nothing Then GoTo CleanUp

    lastRow = wsNKC.Cells(wsNKC.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then GoTo CleanUp

    lastCol = wsNKC.Cells(2, wsNKC.Columns.Count).End(xlToLeft).Column
    Set rng = wsNKC.Range(wsNKC.Cells(2, 1), wsNKC.Cells(lastRow, lastCol))

    ' CLEAR FILTER CU (nhung GIU AutoFilter mode de giu dropdown UI va SUBTOTAL)
    On Error Resume Next
    If wsNKC.FilterMode Then wsNKC.ShowAllData
    wsNKC.Rows.Hidden = False  ' Unhide tat ca rows truoc
    On Error GoTo CleanUp

    ' Bat AutoFilter neu chua co
    If Not wsNKC.AutoFilterMode Then rng.AutoFilter

    ' Cot TK No/Co luon o H/I (8/9) theo yeu cau, cac cot khac van tim theo header
    fieldNo = 8
    fieldCo = 9
    amountField = FindColumnByHeader(wsNKC, Array("so tien", "tien"), 10)
    monthField = FindColumnByHeader(wsNKC, Array("thang", "thang ps", "month"), 3)

    ' Map theo cot click: col C (N?) => tkRoot o TK N?, doi ung o TK Có; col D (Có) => tkRoot o TK Có, doi ung o TK N?
    If isDebitSide Then
        rootField = fieldNo   ' tkRoot o TK N?
        oppField = fieldCo    ' doi ung o TK Có
    Else
        rootField = fieldCo   ' tkRoot o TK Có
        oppField = fieldNo    ' doi ung o TK N?
    End If

    ' Xac dinh prefix cho TK doi ung
    If oppLenSetting > 0 And Len(oppKey) >= oppLenSetting Then
        oppPrefix = Left$(oppKey, oppLenSetting)
    Else
        oppPrefix = oppKey
    End If

    ' Build array va apply filter
    Dim rootArray() As Variant
    Dim oppArray() As Variant
    rootArray = BuildFilterArray(wsNKC, rootField, tkRoot, lastRow)
    oppArray = BuildFilterArray(wsNKC, oppField, oppPrefix, lastRow)

    ' Apply filter with xlFilterValues (an toan loi -> fallback wildcard)
    On Error Resume Next
    If HasFilterValues(rootArray) Then
        rng.AutoFilter Field:=rootField, Criteria1:=rootArray, Operator:=xlFilterValues
        If Err.Number = 0 Then appliedRoot = True
    End If
    Err.Clear

    If HasFilterValues(oppArray) Then
        rng.AutoFilter Field:=oppField, Criteria1:=oppArray, Operator:=xlFilterValues
        If Err.Number = 0 Then appliedOpp = True
    End If
    Err.Clear

    ' Neu auto-filter 0 dong -> thu lai bang criteria text-only
    visCount = VisibleDataRows(rng)
    If visCount = 0 Then
        Dim rootArrayText() As Variant, oppArrayText() As Variant
        rootArrayText = BuildFilterArrayText(wsNKC, rootField, tkRoot, lastRow)
        oppArrayText = BuildFilterArrayText(wsNKC, oppField, oppPrefix, lastRow)
        ' show all va apply lai
        On Error Resume Next
        wsNKC.ShowAllData
        wsNKC.Rows.Hidden = False
        Err.Clear
        If HasFilterValues(rootArrayText) Then rng.AutoFilter Field:=rootField, Criteria1:=rootArrayText, Operator:=xlFilterValues
        If HasFilterValues(oppArrayText) Then rng.AutoFilter Field:=oppField, Criteria1:=oppArrayText, Operator:=xlFilterValues
        On Error GoTo CleanUp
    Else
        On Error GoTo CleanUp
    End If

    ' Neu khong xay dung duoc criteria array (mix number/text), fallback sang wildcard prefix
    If Not appliedRoot And tkRoot <> "" Then
        On Error Resume Next
        rng.AutoFilter Field:=rootField, Criteria1:="=" & tkRoot & "*"
        On Error GoTo CleanUp
    End If

    If Not appliedOpp And oppPrefix <> "" Then
        On Error Resume Next
        rng.AutoFilter Field:=oppField, Criteria1:="=" & oppPrefix & "*"
        On Error GoTo CleanUp
    End If

    On Error Resume Next
    rng.AutoFilter Field:=amountField, Criteria1:=">=0"
    If Err.Number <> 0 Then Err.Clear
    If hasMonth Then rng.AutoFilter Field:=monthField, Criteria1:=CLng(monthVal)
    On Error GoTo CleanUp

CleanUp:
    ' Bat lai screen updating va events
    Application.ScreenUpdating = oldScreen
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents

    ' Activate sheet va select cell SAU KHI bat lai ScreenUpdating
    On Error Resume Next
    wsNKC.Activate
    wsNKC.Range("A3").Select
    On Error GoTo 0
End Sub
Private Function HasFilterValues(arr As Variant) As Boolean
    On Error Resume Next
    HasFilterValues = IsArray(arr)
    If HasFilterValues Then
        HasFilterValues = (UBound(arr) >= LBound(arr)) And Not IsEmpty(arr(LBound(arr)))
    End If
    On Error GoTo 0
End Function
Private Sub EnsureMonthValidation(rng As Range)
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="1,2,3,4,5,6,7,8,9,10,11,12"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub
Private Sub EnsureBoolValidation(rng As Range)
    Dim c As Range
    If rng Is Nothing Then Exit Sub
    For Each c In rng.Cells
        With c.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="TRUE,FALSE"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Next c
End Sub
Private Function BuildFilterArrayText(ws As Worksheet, fieldNum As Long, prefix As String, lastRow As Long) As Variant()
    Dim dict As Object
    Dim arrData As Variant
    Dim i As Long
    Dim tkVal As String
    Dim prefixLen As Long
    Dim result() As Variant
    Dim items As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    prefixLen = Len(prefix)
    If prefixLen = 0 Then
        ReDim result(0 To 0)
        result(0) = Empty
        BuildFilterArrayText = result
        Exit Function
    End If

    If lastRow > 2 Then
        arrData = ws.Range(ws.Cells(3, fieldNum), ws.Cells(lastRow, fieldNum)).Value
        For i = 1 To UBound(arrData, 1)
            If Not IsEmpty(arrData(i, 1)) Then
                tkVal = Trim$(CStr(arrData(i, 1)))
                If tkVal <> "" Then
                    If Left$(tkVal, 1) = "'" Then tkVal = Mid$(tkVal, 2)
                    If Len(tkVal) >= prefixLen Then
                        If Left$(tkVal, prefixLen) = prefix Then
                            If Not dict.Exists(tkVal) Then dict.Add tkVal, tkVal
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If dict.Count > 0 Then
        items = dict.Items
        ReDim result(1 To dict.Count)
        For i = 0 To UBound(items)
            result(i + 1) = items(i)
        Next i
    Else
        ReDim result(0 To 0)
        result(0) = Empty
    End If

    BuildFilterArrayText = result
End Function
Private Function VisibleDataRows(rng As Range) As Long
    Dim vis As Range
    On Error Resume Next
    Set vis = rng.Offset(1, 0).Resize(rng.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If vis Is Nothing Then
        VisibleDataRows = 0
    Else
        VisibleDataRows = vis.Rows.Count
    End If
End Function
Private Function FindColumnByHeader(ws As Worksheet, names As Variant, defaultCol As Long) As Long
    Dim maxCol As Long, r As Long, c As Long
    Dim hdr As String, target As Variant
    On Error Resume Next
    maxCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If maxCol < 1 Then
        FindColumnByHeader = defaultCol
        Exit Function
    End If
    For r = 1 To 3
        For c = 1 To maxCol
            hdr = NormalizeHeader(ws.Cells(r, c).Value)
            For Each target In names
                If hdr Like NormalizeHeader(target) Then
                    FindColumnByHeader = c
                    Exit Function
                End If
            Next target
        Next c
    Next r
    FindColumnByHeader = defaultCol
End Function
Private Function NormalizeHeader(v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then NormalizeHeader = "": Exit Function
    s = LCase$(Trim$(CStr(v)))
    s = RemoveAccentsAscii(s)
    NormalizeHeader = s
End Function
Private Function RemoveAccentsAscii(text As String) As String
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
    RemoveAccentsAscii = text
End Function

