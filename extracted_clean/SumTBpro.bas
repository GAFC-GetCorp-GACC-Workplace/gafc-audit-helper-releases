Attribute VB_Name = "SumTBpro"
Option Explicit
Public Sub Sumpro(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsTB As Worksheet, wsXL As Worksheet
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsTB = RequireSheet(wb, "TB", "Chua co sheet 'TB'. Hay tao TB truoc.")
    If wsTB Is Nothing Then Exit Sub
    Set wsXL = RequireSheet(wb, "Xu_ly", "Chua co sheet 'Xu_ly'. Hay chay TaoTB truoc.")
    If wsXL Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim lastTB As Long: lastTB = wsTB.Cells(wsTB.Rows.Count, 3).End(xlUp).Row
    Dim lastXL As Long: lastXL = wsXL.Cells(wsXL.Rows.Count, 1).End(xlUp).Row
    Dim dataXL As Variant
    dataXL = wsXL.Range("A2:G" & lastXL).Value
    Dim i As Long, r As Long, c As Long
    Dim prefix As String, tk As String, tk3 As String
    Dim resultArr() As Variant
    ReDim resultArr(1 To lastTB - 1, 1 To 6)
    Dim specialTKs As Variant
    specialTKs = Array("2141", "2142", "2143", "4211", "4212", "2147", "2421", _
                       "2422", "8211", "8212", "2441", "2442", "3411", "3412", "2444", "2291", "2292", "2293", "2294")
    Dim dictFull As Object, dictPrefix3 As Object
    Dim dictSpecial As Object, dictSpecialPrefix As Object
    Dim sums As Variant
    Dim k As Variant
    Set dictFull = CreateObject("Scripting.Dictionary")
    Set dictPrefix3 = CreateObject("Scripting.Dictionary")
    Set dictSpecial = CreateObject("Scripting.Dictionary")
    Set dictSpecialPrefix = CreateObject("Scripting.Dictionary")
    For Each k In specialTKs
        dictSpecial(CStr(k)) = True
    Next k

    ' Precompute sums from Xu_ly once
    For r = 1 To UBound(dataXL, 1)
        If IsError(dataXL(r, 1)) Or IsEmpty(dataXL(r, 1)) Then GoTo NextXLRow
        tk = Trim$(CStr(dataXL(r, 1)))
        If tk <> "" Then
            If Not dictFull.Exists(tk) Then dictFull.Add tk, Array(0#, 0#, 0#, 0#, 0#, 0#)
            sums = dictFull(tk)
            For c = 0 To 5
                If IsNumeric(dataXL(r, c + 2)) Then
                    sums(c) = sums(c) + CDbl(dataXL(r, c + 2))
                End If
            Next c
            dictFull(tk) = sums

            If Len(tk) >= 3 Then
                tk3 = Left$(tk, 3)
                If Not dictPrefix3.Exists(tk3) Then dictPrefix3.Add tk3, Array(0#, 0#, 0#, 0#, 0#, 0#)
                sums = dictPrefix3(tk3)
                For c = 0 To 5
                    If IsNumeric(dataXL(r, c + 2)) Then
                        sums(c) = sums(c) + CDbl(dataXL(r, c + 2))
                    End If
                Next c
                dictPrefix3(tk3) = sums
            End If

            For Each k In specialTKs
                If Left$(tk, Len(CStr(k))) = CStr(k) Then
                    If Not dictSpecialPrefix.Exists(CStr(k)) Then dictSpecialPrefix.Add CStr(k), Array(0#, 0#, 0#, 0#, 0#, 0#)
                    sums = dictSpecialPrefix(CStr(k))
                    For c = 0 To 5
                        If IsNumeric(dataXL(r, c + 2)) Then
                            sums(c) = sums(c) + CDbl(dataXL(r, c + 2))
                        End If
                    Next c
                    dictSpecialPrefix(CStr(k)) = sums
                End If
            Next k
        End If
NextXLRow:
    Next r

    For i = 2 To lastTB
        If IsError(wsTB.Cells(i, 3).Value) Then GoTo NextTBRow
        prefix = Trim$(CStr(wsTB.Cells(i, 3).Value))
        If prefix <> "" Then
            If dictSpecial.Exists(prefix) Then
                If dictFull.Exists(prefix) Then
                    sums = dictFull(prefix)
                ElseIf dictSpecialPrefix.Exists(prefix) Then
                    sums = dictSpecialPrefix(prefix)
                Else
                    sums = Array(0#, 0#, 0#, 0#, 0#, 0#)
                End If
            Else
                If dictFull.Exists(prefix) Then
                    sums = dictFull(prefix)
                ElseIf Len(prefix) >= 3 Then
                    tk3 = Left$(prefix, 3)
                    If dictPrefix3.Exists(tk3) Then
                        sums = dictPrefix3(tk3)
                    Else
                        sums = Array(0#, 0#, 0#, 0#, 0#, 0#)
                    End If
                Else
                    sums = Array(0#, 0#, 0#, 0#, 0#, 0#)
                End If
            End If
            For c = 1 To 6
                resultArr(i - 1, c) = Round(CDbl(sums(c - 1)), 0)
            Next c
        Else
            For c = 1 To 6
                resultArr(i - 1, c) = 0#
            Next c
        End If
NextTBRow:
    Next i
    wsTB.Range("D2").Resize(UBound(resultArr), 6).Value = resultArr
    Dim totalRow As Long: totalRow = lastTB + 1
    wsTB.Cells(totalRow, 3).Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    Dim colLetter As Variant: colLetter = Array("D", "E", "F", "G", "H", "I")
    For c = 0 To 5
        wsTB.Cells(totalRow, c + 4).Formula = "=SUM(" & colLetter(c) & "2:" & colLetter(c) & lastTB & ")"
    Next c
    With wsTB.Range(wsTB.Cells(totalRow, 3), wsTB.Cells(totalRow, 9))
        .Font.Bold = True
        .Interior.Color = RGB(150, 240, 240)
    End With
    ' === NEW FUNCTIONALITY: Difference Row ===
    Dim diffRow As Long: diffRow = totalRow + 1
    wsTB.Cells(diffRow, 3).Value = "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch"
    wsTB.Cells(diffRow, 4).Formula = "=" & colLetter(0) & totalRow & "-" & colLetter(1) & totalRow
    wsTB.Cells(diffRow, 6).Formula = "=" & colLetter(2) & totalRow & "-" & colLetter(3) & totalRow
    wsTB.Cells(diffRow, 8).Formula = "=" & colLetter(4) & totalRow & "-" & colLetter(5) & totalRow
    With wsTB.Range(wsTB.Cells(diffRow, 3), wsTB.Cells(diffRow, 9))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    ' === NEW FUNCTIONALITY: Format Cells ===
    With wsTB.Range("D2:I" & diffRow)
        .NumberFormat = "#,##0"
    End With
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    InfoToast "Done"
End Sub

