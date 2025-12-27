Attribute VB_Name = "Chuyen_dau_ki"
Option Explicit
Public Sub chuyendauki(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not ConfirmActiveSheetRisk("Hay dam bao dung sheet mau truoc khi chay.") Then Exit Sub
    ' --- Ngu?i dùng nh?p thông tin ---
    Dim colFrom As String, colTo As String, colClear As String
    colFrom = UCase(InputBox("Nhap cot cuoi ki(nhap - neu chi muon xoa data):", "Cot cuoi ki"))
    ' N?u ngu?i dùng ch? mu?n xóa
    If colFrom = "-" Then
        colClear = UCase(InputBox("Nhap cot can xoa:", "Xoa cot"))
        If colClear = "" Or colClear = "-" Then
            MsgBox "Ban chua nhap cot can xoa!", vbExclamation
            Exit Sub
        End If
        Call OnlyClearData(ws, colClear)
        Exit Sub
    End If
    ' N?u ngu?i dùng mu?n copy (có th? có ho?c không xóa)
    colTo = UCase(InputBox("Nhap cot dau ki", "Cot dau ki"))
    If colFrom = "" Or colTo = "" Then
        MsgBox "Ban chua nhap cot dau ki", vbExclamation
        Exit Sub
    End If
    colClear = UCase(InputBox("Nhap cot can xoa:", "Xoa cot"))
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' === COPY D? LI?U ===
    Dim copyRanges As Variant
    copyRanges = Array("11:14", "18:21", "23:25", "28:37", "40:48", _
                       "52:57", "62:69", "73:74", "76:77", "79:80", _
                       "83:84", "87:88", "91:95", "99:103", "109:122", _
                       "125:137", "143:145", "146:154", "156:157", _
                       "158:159", "161:162", "170:170", "173:176", _
                       "180:180", "184:191", "195:196", "201:202")
    Dim i As Long
    For i = LBound(copyRanges) To UBound(copyRanges)
        ws.Range(colTo & Split(copyRanges(i), ":")(0) & ":" & colTo & Split(copyRanges(i), ":")(1)).Value = _
            ws.Range(colFrom & Split(copyRanges(i), ":")(0) & ":" & colFrom & Split(copyRanges(i), ":")(1)).Value
    Next i
    ' === CÔNG TH?C ===
    ws.Range(colTo & "10").Formula = "=SUM(" & colTo & "11:" & colTo & "14)"
    ws.Range(colTo & "17").Formula = "=SUM(" & colTo & "18:" & colTo & "21)"
    ws.Range(colTo & "22").Formula = "=SUM(" & colTo & "23:" & colTo & "25)"
    ws.Range(colTo & "16").Formula = "=" & colTo & "17+" & colTo & "22"
    ws.Range(colTo & "27").Formula = "=SUM(" & colTo & "28:" & colTo & "37)"
    ws.Range(colTo & "39").Formula = "=SUM(" & colTo & "40:" & colTo & "48)"
    ws.Range(colTo & "50").Formula = "=SUM(" & colTo & "52:" & colTo & "57)"
    ws.Range(colTo & "61").Formula = "=SUM(" & colTo & "62:" & colTo & "69)"
    ws.Range(colTo & "72").Formula = "=SUM(" & colTo & "73:" & colTo & "74)"
    ws.Range(colTo & "75").Formula = "=SUM(" & colTo & "76:" & colTo & "77)"
    ws.Range(colTo & "78").Formula = "=SUM(" & colTo & "79:" & colTo & "80)"
    ws.Range(colTo & "82").Formula = "=SUM(" & colTo & "83:" & colTo & "84)"
    ws.Range(colTo & "86").Formula = "=SUM(" & colTo & "87:" & colTo & "88)"
    ws.Range(colTo & "90").Formula = "=SUM(" & colTo & "91:" & colTo & "95)"
    ws.Range(colTo & "98").Formula = "=SUM(" & colTo & "99:" & colTo & "103)"
    ws.Range(colTo & "108").Formula = "=SUM(" & colTo & "109:" & colTo & "122)"
    ws.Range(colTo & "124").Formula = "=SUM(" & colTo & "125:" & colTo & "137)"
    ws.Range(colTo & "142").Formula = "=SUM(" & colTo & "143:" & colTo & "145)"
    ws.Range(colTo & "155").Formula = "=SUM(" & colTo & "156:" & colTo & "157)"
    ws.Range(colTo & "160").Formula = "=SUM(" & colTo & "161:" & colTo & "162)"
    ws.Range(colTo & "172").Formula = "=SUM(" & colTo & "173:" & colTo & "176)"
    ws.Range(colTo & "60").Formula = "=" & colTo & "61+" & colTo & "71+" & colTo & "82+" & colTo & "86+" & colTo & "90+" & colTo & "98"
    ' === XOÁ D? LI?U (n?u ngu?i dùng mu?n) ===
    If colClear <> "-" And colClear <> "" Then
        Call OnlyClearData(ws, colClear)
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    InfoToast "Done" & vbCrLf & _
           "?? Copy t?: " & colFrom & " ? " & colTo & vbCrLf & _
           IIf(colClear = "-" Or colClear = "", "? Không th?c hi?n xóa.", "?? Ðã xóa d? li?u t?i c?t " & colClear), vbInformation
End Sub
' === HÀM PH? ===
Private Sub OnlyClearData(ws As Worksheet, colClear As String)
    Dim clearRanges As Variant, i As Long
    clearRanges = Array("11:14", "18:21", "23:25", "28:37", "40:48", _
                        "52:57", "62:69", "73:74", "76:77", "79:80", _
                        "83:84", "87:88", "91:95", "99:103", "109:122", _
                        "125:137", "143:154", "156:157", "158:159", _
                        "161:162", "170:170", "173:176", "180:180", _
                        "184:191", "195:196", "201:202")
    For i = LBound(clearRanges) To UBound(clearRanges)
        ws.Range(colClear & Split(clearRanges(i), ":")(0) & ":" & colClear & Split(clearRanges(i), ":")(1)).ClearContents
    Next i
    InfoToast "Done, da xoa xong " & colClear
End Sub

