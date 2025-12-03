Option Explicit
'==== Ð?nh nghia c?t trong detailData ====
Private Const COL_SO_CT  As Long = 1   ' Col A = S? CT
Private Const COL_NGAY   As Long = 3   ' Col C = Ngày
Private Const COL_MA     As Long = 4   ' Col D = Mã hàng
Private Const COL_SL     As Long = 16  ' Col P = S? lu?ng
Private Const COL_GIATRI As Long = 17  ' Col Q = Giá tr?
'==============================================================
' Th? t?c chính
'==============================================================
Public Sub detailtopricing(control As IRibbonControl)
'=================================================================
' M?c dích:
'  - Copy d? li?u t? DETAIL sang PRICING
'  - B? qua dòng có SL=0 ho?c Giá tr?=0
'  - Gom hóa don cho d?n khi d? s? lu?ng
'  - Thêm dòng Total NGAY SAU HEADER
'  - Dòng Total có thêm c?t M = (Ðon giá Total - Ðon giá g?c)/Ðon giá g?c
'=================================================================
    Dim wsDetail As Worksheet, wsTarget As Worksheet
    Dim lastRowDetail As Long, lastRowTarget As Long
    Dim detailData As Variant, targetData As Variant
    Dim i As Long, j As Long, k As Long
    Dim codeTarget As String
    Dim qtyNeeded As Double, qtyAccum As Double
    Dim resultRow As Long
    Dim filterOn As Boolean
    Dim dict As Object
    Dim items As Collection
    Dim rec As Variant
    Dim calcMode As XlCalculation
    On Error GoTo CleanFail
    '--- Tang t?c Application ---
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        calcMode = .Calculation
        .Calculation = xlCalculationManual
    End With
    '--- Gán sheet ---
    Set wsDetail = Worksheets("D550.1.1 Detail Input")
    Set wsTarget = Worksheets("D550.1 Pricing Testing RW-M")
    '--- T?t filter n?u có ---
    filterOn = wsDetail.AutoFilterMode
    If filterOn Then wsDetail.AutoFilterMode = False
    '--- L?y vùng d? li?u ---
    lastRowDetail = wsDetail.Cells(wsDetail.Rows.Count, "A").End(xlUp).Row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    If lastRowDetail < 3 Or lastRowTarget < 3 Then GoTo SafeExit
    detailData = wsDetail.Range("A3:Q" & lastRowDetail).Value2
    targetData = wsTarget.Range("A3:G" & lastRowTarget).Value2
    '--- Xoá d? li?u cu ---
    wsTarget.Range("A3:M" & wsTarget.Rows.Count).ClearContents
    resultRow = 3
    '==============================================================
    ' 1) T?o index các dòng Detail theo Mã hàng (quét ngu?c)
    '==============================================================
    Set dict = CreateObject("Scripting.Dictionary")
    Dim rowCnt As Long: rowCnt = UBound(detailData, 1)
    For j = rowCnt To 1 Step -1
        Dim code As String
        code = Trim$(CStr(detailData(j, COL_MA)))
        If Len(code) > 0 Then
            If IsNumeric(detailData(j, COL_SL)) And IsNumeric(detailData(j, COL_GIATRI)) Then
                If CDbl(detailData(j, COL_SL)) > 0 And CDbl(detailData(j, COL_GIATRI)) > 0 Then
                    Dim dt As Variant
                    dt = SafeDate(detailData(j, COL_NGAY))
                    If Not IsEmpty(dt) Then
                        If Not dict.Exists(code) Then
                            Set dict(code) = New Collection
                        End If
                        rec = Array( _
                            dt, _
                            detailData(j, COL_SO_CT), _
                            CDbl(detailData(j, COL_SL)), _
                            CDbl(detailData(j, COL_GIATRI)), _
                            j _
                        )
                        dict(code).Add rec
                    End If
                End If
            End If
        End If
    Next j
    '==============================================================
    ' 2) Duy?t t?ng mã hàng trong Pricing & ghi k?t qu?
    '==============================================================
    Dim startRow As Long, endRow As Long
    Dim gomCount As Long
    For i = 1 To UBound(targetData, 1)
        If Len(Trim$(CStr(targetData(i, 2)))) > 0 Then
            codeTarget = Trim$(CStr(targetData(i, 2)))
            qtyNeeded = NzDbl(targetData(i, 5))
            qtyAccum = 0
            If dict.Exists(codeTarget) Then
                Set items = dict(codeTarget)
                If items.Count > 0 Then
                    '--- Ghi 7 c?t d?u ---
                    wsTarget.Cells(resultRow, 1).Resize(1, 7).Value = SliceRow(targetData, i, 1, 7)
                    wsTarget.Cells(resultRow, 7).FormulaR1C1 = "=RC[-3]/RC[-2]"
                    startRow = resultRow + 1   ' dòng Total s? n?m ? dây
                    ' b? tr?ng dòng Total, ghi chi ti?t t? dòng sau
                    resultRow = resultRow + 2
                    '--- Gom hoá don ---
                    gomCount = 0
                    For k = 1 To items.Count
                        rec = items(k)
                        qtyAccum = qtyAccum + NzDbl(rec(2))
                        gomCount = gomCount + 1
                        wsTarget.Cells(resultRow, 8).Value = rec(0)
                        wsTarget.Cells(resultRow, 9).Value = rec(1)
                        wsTarget.Cells(resultRow, 10).Value = rec(2)
                        wsTarget.Cells(resultRow, 11).Value = rec(3)
                        wsTarget.Cells(resultRow, 12).FormulaR1C1 = "=RC[-1]/RC[-2]"
                        resultRow = resultRow + 1
                        If qtyAccum >= qtyNeeded Then Exit For
                    Next k
                    '--- Dòng Total ---
                    endRow = resultRow - 1   ' last detail row
                    wsTarget.Cells(startRow, 9).Value = "Total"
                    wsTarget.Cells(startRow, 10).FormulaR1C1 = "=SUM(R[1]C:R[" & (endRow - startRow) & "]C)"
                    wsTarget.Cells(startRow, 11).FormulaR1C1 = "=SUM(R[1]C:R[" & (endRow - startRow) & "]C)"
                    wsTarget.Cells(startRow, 12).FormulaR1C1 = "=RC[-1]/RC[-2]"
                    With wsTarget.Range(wsTarget.Cells(startRow, 9), wsTarget.Cells(startRow, 12))
                        .Interior.Color = RGB(200, 255, 200)
                        .Font.Bold = True
                    End With
                    ' --- Thêm c?t M: so sánh don giá ---
                    wsTarget.Cells(startRow, 13).FormulaR1C1 = "=(R[-" & (startRow - (startRow - 1)) & "]C[-6]-RC[-1])/RC[-1]"
                    wsTarget.Cells(startRow, 13).NumberFormat = "0.00%"
                End If
            Else
                '--- Không tìm th?y mã ---
                wsTarget.Cells(resultRow, 1).Resize(1, 7).Value = SliceRow(targetData, i, 1, 7)
                With wsTarget.Cells(resultRow, 8)
                    .Value = "Không tìm th?y mã hàng"
                    .Interior.Color = RGB(255, 255, 150)
                End With
                resultRow = resultRow + 1
            End If
        End If
    Next i
SafeExit:
    '--- Khôi ph?c filter ---
    If filterOn Then wsDetail.Rows("2:2").AutoFilter
    '--- Format k?t qu? ---
    With wsTarget
        If resultRow > 3 Then
            .Range("J3:L" & resultRow - 1).NumberFormat = "#,##0"
            .Range("H3:H" & resultRow - 1).NumberFormat = "dd/mm/yyyy"
        End If
        .Columns("A:M").AutoFit
    End With
    MsgBox "Done", vbInformation
CleanFinally:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calcMode
    End With
    Exit Sub
CleanFail:
    MsgBox ChrW(272) & ChrW(227) & " x" & ChrW(7843) & "y ra l" & ChrW(7895) & "i" - " & Err.Description, vbExclamation"
    Resume CleanFinally
End Sub
'==============================================================
' Helpers
'==============================================================
' Parse ngày an toàn
Private Function SafeDate(ByVal v As Variant) As Variant
    On Error GoTo Bad
    If IsDate(v) Then
        SafeDate = CDate(v): Exit Function
    End If
    If IsNumeric(v) Then
        SafeDate = CDate(CDbl(v)): Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then GoTo Bad
    Dim sep As String
    If InStr(s, "/") > 0 Then sep = "/" Else sep = "-"
    Dim t() As String: t = Split(s, sep)
    If UBound(t) <> 2 Then GoTo Bad
    Dim d As Long, m As Long, y As Long
    d = val(t(0)): m = val(t(1)): y = val(t(2))
    If y < 100 Then y = 2000 + y
    SafeDate = DateSerial(y, m, d)
    Exit Function
Bad:
    SafeDate = Empty
End Function
' Trích 1 hàng thành m?ng 1xN
Private Function SliceRow(ByRef arr As Variant, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As Variant
    Dim n As Long: n = c2 - c1 + 1
    Dim tmp() As Variant
    ReDim tmp(1 To 1, 1 To n)
    Dim c As Long
    For c = 1 To n
        tmp(1, c) = arr(r, c1 + c - 1)
    Next c
    SliceRow = tmp
End Function
' Null/Empty -> 0
Private Function NzDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        NzDbl = 0#
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0#
    End If
End Function