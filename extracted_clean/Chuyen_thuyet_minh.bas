Option Explicit
Public Sub chuyen_tm11(control As IRibbonControl)
    Dim ws As Worksheet
    Dim firstRow As Long, lastRow As Long
    Dim rowIndex As Long
    Dim label As String
    Dim f As String
    Set ws = ActiveSheet ' Dùng sheet hi?n t?i
    ' Nh?p vùng dòng x? lý
    firstRow = Application.InputBox("Nh?p s? b?t d?u", Type:=1)
    If firstRow = 0 Then Exit Sub
    lastRow = Application.InputBox("Nh?p s? k?t thúc", Type:=1)
    If lastRow = 0 Then Exit Sub
    If lastRow < firstRow Then
        MsgBox "Dòng k?t thúc ph?i l?n hon dòng b?t d?u!", vbExclamation
        Exit Sub
    End If
    ' Tang t?c
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' L?p qua t?ng dòng
    For rowIndex = firstRow To lastRow
        label = ws.Cells(rowIndex, "B").Text
        If IsTongCong(label) Then
            ' Dòng T?ng c?ng
            If IsActualFormula(ws.Cells(rowIndex, "L")) Then
                If IsTextValue(ws.Cells(rowIndex, "L")) Then
                    ' Không copy n?u là text
                ElseIf HasExternalSheetReference(ws.Cells(rowIndex, "L")) Then
                    ws.Cells(rowIndex, "P").Value = ws.Cells(rowIndex, "L").Value
                Else
                    f = ws.Cells(rowIndex, "L").Formula
                    ws.Cells(rowIndex, "P").Formula = ConvertMergedColumn(f, "L", "M", "N", "P", "Q", "R")
                End If
            ElseIf IsRealNumber(ws.Cells(rowIndex, "L")) Then
                ws.Cells(rowIndex, "P").Value = ws.Cells(rowIndex, "L").Value
            End If
        Else
            ' Dòng thu?ng
            If IsActualFormula(ws.Cells(rowIndex, "L")) Then
                If IsTextValue(ws.Cells(rowIndex, "L")) Then
                    ' Không copy n?u là text
                ElseIf HasExternalSheetReference(ws.Cells(rowIndex, "L")) Then
                    ws.Cells(rowIndex, "P").Value = ws.Cells(rowIndex, "L").Value
                Else
                    f = ws.Cells(rowIndex, "L").Formula
                    ws.Cells(rowIndex, "P").Formula = ConvertMergedColumn(f, "L", "M", "N", "P", "Q", "R")
                End If
            ElseIf IsRealNumber(ws.Cells(rowIndex, "L")) And Not IsTitleText(ws.Cells(rowIndex, "L").Text) Then
                ws.Cells(rowIndex, "P").Value = ws.Cells(rowIndex, "L").Value
            End If
        End If
    Next rowIndex
    ' Khôi ph?c Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Hoàn thành", vbInformation
End Sub
' ===========================================
' Hàm nh?n di?n "T?ng c?ng"
' ===========================================
Function IsTongCong(Txt As String) As Boolean
    Dim cleaned As String
    cleaned = Trim(Txt)
    ' Cách 1: d? dài = 9 và ký t? cu?i là "g"
    If Len(cleaned) = 9 Then
        If LCase(Right(cleaned, 1)) = "g" Then
            IsTongCong = True
            Exit Function
        End If
    End If
    ' Cách 2: ki?m tra tr?c ti?p
    cleaned = LCase(Replace(cleaned, " ", ""))
    If cleaned = "tongcong" Or cleaned = "t?ngc?ng" Then
        IsTongCong = True
        Exit Function
    End If
    IsTongCong = False
End Function
' ===========================================
' Ki?m tra lo?i d? li?u ô
' ===========================================
Function IsRealNumber(cell As Range) As Boolean
    On Error Resume Next
    IsRealNumber = Not IsEmpty(cell.Value) And IsNumeric(cell.Value) And Not IsDate(cell.Value)
    On Error GoTo 0
End Function
Function IsTextValue(cell As Range) As Boolean
    If IsEmpty(cell.Value) Then
        IsTextValue = False
    ElseIf IsNumeric(cell.Value) Then
        IsTextValue = False
    Else
        IsTextValue = True
    End If
End Function
Function IsTitleText(Txt As String) As Boolean
    Dim t As String
    t = LCase(Trim(Txt))
    IsTitleText = (t = "nam nay" Or t = "nam tru?c" Or t = "s? cu?i nam" Or t = "s? d?u nam")
End Function
Function IsActualFormula(cell As Range) As Boolean
    On Error Resume Next
    Dim f As String
    f = Trim(cell.Formula)
    IsActualFormula = (Left(f, 1) = "=")
    On Error GoTo 0
End Function
Function HasExternalSheetReference(cell As Range) As Boolean
    Dim f As String
    f = cell.Formula
    If InStr(f, "!") > 0 Then
        HasExternalSheetReference = True
        Exit Function
    End If
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    ' Thêm c?t P vào pattern: ki?m tra c? T và P
    regex.Pattern = "(\$?[TP]\$?\d+|[TP]:[TP])"
    regex.Global = False
    HasExternalSheetReference = regex.Test(f)
End Function
' ===========================================
' Hàm chuy?n công th?c
' ===========================================
Function ConvertMergedColumn(formulaText As String, col1 As String, col2 As String, col3 As String, toCol1 As String, toCol2 As String, toCol3 As String) As String
    Dim result As String
    result = formulaText
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' N?u là SUM thì d?c bi?t: SUM(L..N) ? SUM(P..R)
    regex.Pattern = "SUM\s*\(\s*" & col1 & "(\d+)\s*:\s*" & col3 & "(\d+)\s*\)"
    If regex.Test(result) Then
        result = regex.Replace(result, "SUM(" & toCol1 & "$1:" & toCol3 & "$2)")
    Else
        ' N?u không ph?i SUM thì thay bình thu?ng
        result = ConvertNormalFormula(result, col1, col2, col3, toCol1, toCol2, toCol3)
    End If
    ConvertMergedColumn = result
End Function
Function ConvertNormalFormula(formulaText As String, col1 As String, col2 As String, col3 As String, toCol1 As String, toCol2 As String, toCol3 As String) As String
    Dim result As String
    result = formulaText
    ' Thay t?ng c?t
    result = Replace(result, col1, toCol1)
    result = Replace(result, col2, toCol2)
    result = Replace(result, col3, toCol3)
    ConvertNormalFormula = result
End Function