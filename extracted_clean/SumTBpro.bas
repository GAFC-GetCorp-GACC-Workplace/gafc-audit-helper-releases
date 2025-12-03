Option Explicit
Public Sub Sumpro(control As IRibbonControl)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim wsTB As Worksheet, wsXL As Worksheet
    Set wsTB = Worksheets("TB")
    Set wsXL = Worksheets("Xu_ly")
    Dim lastTB As Long: lastTB = wsTB.Cells(wsTB.Rows.Count, 3).End(xlUp).Row
    Dim lastXL As Long: lastXL = wsXL.Cells(wsXL.Rows.Count, 1).End(xlUp).Row
    Dim dataXL As Variant
    dataXL = wsXL.Range("A2:G" & lastXL).Value
    Dim i As Long, r As Long, c As Long
    Dim prefix As String, tk As String
    Dim sumArr(1 To 6) As Double
    Dim resultArr() As Variant
    ReDim resultArr(1 To lastTB - 1, 1 To 6)
    Dim specialTKs As Variant
    specialTKs = Array("2141", "2142", "2143", "4211", "4212", "2147", "2421", _
                       "2422", "8211", "8212", "2441", "2442", "3411", "3412", "2444", "2291", "2292", "2293", "2294")
    Dim hasExact As Boolean
    For i = 2 To lastTB
        prefix = Trim(wsTB.Cells(i, 3).Value)
        For c = 1 To 6: sumArr(c) = 0#: Next c
        hasExact = False
        For r = 1 To UBound(dataXL, 1)
            If Trim(CStr(dataXL(r, 1))) = prefix Then
                hasExact = True
                Exit For
            End If
        Next r
        For r = 1 To UBound(dataXL, 1)
            tk = Trim(CStr(dataXL(r, 1)))
            If Not IsError(Application.match(prefix, specialTKs, 0)) Then
                If tk = prefix Or (Not hasExact And Left(tk, Len(prefix)) = prefix And tk <> prefix) Then
                    For c = 1 To 6
                        If IsNumeric(dataXL(r, c + 1)) Then
                            sumArr(c) = sumArr(c) + dataXL(r, c + 1)
                        End If
                    Next c
                End If
            Else
                If (hasExact And tk = prefix) Or (Not hasExact And Left(tk, 3) = prefix) Then
                    For c = 1 To 6
                        If IsNumeric(dataXL(r, c + 1)) Then
                            sumArr(c) = sumArr(c) + dataXL(r, c + 1)
                        End If
                    Next c
                End If
            End If
        Next r
        ' THAY Ð?I ? ÐÂY - Làm tròn theo quy t?c chu?n
        For c = 1 To 6
            resultArr(i - 1, c) = Round(sumArr(c), 0)  ' Làm tròn chu?n qu?c t?
        Next c
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
    MsgBox "Done", vbInformation
End Sub