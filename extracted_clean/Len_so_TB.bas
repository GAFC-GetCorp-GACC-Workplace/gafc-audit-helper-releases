Attribute VB_Name = "Len_so_TB"
Option Explicit
Public Sub Len_so_bao_cao(control As IRibbonControl)
    Dim wb As Workbook
    Dim wsTB As Worksheet, wsFS As Worksheet
    Dim lastRowTB As Long, lastRowFS As Long
    Dim i As Long, j As Long
    Dim code1 As String, code2 As String
    Dim val1 As Double, val2 As Double
    Dim arrFS As Variant
    Dim tempVal2 As Double
    ' ?? L?y workbook c?a sheet dang active (tránh dính file Add-in)
    Set wb = ActiveSheet.Parent
    ' ?? Ki?m tra t?n t?i 2 sheet
    On Error Resume Next
    Set wsTB = wb.Sheets("TB")
    Set wsFS = wb.Sheets("Adjusted FS")
    On Error GoTo 0
    If wsTB Is Nothing Or wsFS Is Nothing Then
        MsgBox "? Không tìm th?y sheet 'TB' ho?c 'Adjusted FS' trong file dang m?!", vbCritical, "L?i"
        Exit Sub
    End If
    ' ?? Xác d?nh dòng cu?i trong TB
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "H").End(xlUp).Row
    lastRowFS = 250  ' ho?c có th? d?t d?ng theo D c?t n?u b?n mu?n
    ' ?? Ð?c toàn b? c?t D bên Adjusted FS
    arrFS = wsFS.Range("D1:D" & lastRowFS).Value
    ' ?? Duy?t t?ng dòng trong TB
    For i = 2 To lastRowTB
        code1 = Trim(CStr(wsTB.Cells(i, "A").Value))
        code2 = Trim(CStr(wsTB.Cells(i, "B").Value))
        val1 = 0: val2 = 0
        If IsNumeric(wsTB.Cells(i, "H").Value) Then val1 = wsTB.Cells(i, "H").Value
        If IsNumeric(wsTB.Cells(i, "I").Value) Then val2 = wsTB.Cells(i, "I").Value
        ' ?? Tru?ng h?p d?c bi?t: code1 = code2 = 4211 ho?c 4212 ? (-val1 + val2)
        If (code1 = "4211" Or code1 = "4212") And code1 = code2 Then
            For j = 1 To UBound(arrFS, 1)
                If Trim(CStr(arrFS(j, 1))) = code1 Then
                    With wsFS.Cells(j, "G")
                        If IsNumeric(.Value) Then
                            .Value = .Value + (-val1 + val2)
                        Else
                            .Value = -val1 + val2
                        End If
                    End With
                    Exit For
                End If
            Next j
        Else
            ' ?? CODE1: Ghi tr?c ti?p (gi? nguyên d?u)
            If code1 <> "" Then
                For j = 1 To UBound(arrFS, 1)
                    If Trim(CStr(arrFS(j, 1))) = code1 Then
                        With wsFS.Cells(j, "G")
                            If IsNumeric(.Value) Then
                                .Value = .Value + val1
                            Else
                                .Value = val1
                            End If
                        End With
                        Exit For
                    End If
                Next j
            End If
            ' ?? CODE2: M?t s? mã s? ghi âm
            If code2 <> "" Then
                tempVal2 = val2
                If code2 = "2141" Or code2 = "2142" Or code2 = "2143" Or _
                   code2 = "2417" Or code2 = "139" Or code2 = "159" Then
                    tempVal2 = -val2
                End If
                For j = 1 To UBound(arrFS, 1)
                    If Trim(CStr(arrFS(j, 1))) = code2 Then
                        With wsFS.Cells(j, "G")
                            If IsNumeric(.Value) Then
                                .Value = .Value + tempVal2
                            Else
                                .Value = tempVal2
                            End If
                        End With
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    ' ? Thông báo hoàn t?t
    MsgBox "? Ðã chuy?n d? li?u sang Adjusted FS thành công!", vbInformation, "Hoàn t?t"
End Sub

