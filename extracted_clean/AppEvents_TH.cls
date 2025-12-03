Option Explicit
Public WithEvents App As Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ExitHandler
    If Sh Is Nothing Then Exit Sub
    If Sh.Name <> "TH" Then Exit Sub
    ' Cap nhat dropdown C4 khi B4/D3 thay doi
    If Not Intersect(Target, Sh.Range("B4,D3")) Is Nothing Then
        Application.EnableEvents = False
        Update_TH_Dropdown_Internal Sh
        Application.EnableEvents = True
    End If
    ' Refresh TH khi C4, D2, D3 thay doi
    If Intersect(Target, Sh.Range("C4,D2,D3")) Is Nothing Then Exit Sub
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
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If Sh.Name <> "TH" Then Exit Sub
    Refresh_TH Sh.Parent
End Sub