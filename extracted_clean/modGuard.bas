Attribute VB_Name = "modGuard"
Option Explicit

Public Function GetSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function RequireSheet(ByVal wb As Workbook, ByVal sheetName As String, Optional ByVal hint As String = "") As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(wb, sheetName)
    If ws Is Nothing Then
        If hint <> "" Then
            MsgBox hint, vbExclamation
        Else
            MsgBox "Khong tim thay sheet '" & sheetName & "'.", vbExclamation
        End If
    End If
    Set RequireSheet = ws
End Function

Public Function ConfirmProceed(ByVal msg As String) As Boolean
    ConfirmProceed = (MsgBox(msg, vbYesNo + vbQuestion, "Xac nhan") = vbYes)
End Function

Public Function EnsureActiveSheet(ByVal wb As Workbook, ByVal expectedName As String, Optional ByVal warnMsg As String = "") As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    If ActiveSheet Is Nothing Then Exit Function
    If StrComp(ActiveSheet.Name, expectedName, vbTextCompare) = 0 Then
        EnsureActiveSheet = True
        Exit Function
    End If
    If warnMsg = "" Then
        warnMsg = "Macro nay nen chay tren sheet '" & expectedName & "'. Chuyen sang sheet nay khong?"
    End If
    If ConfirmProceed(warnMsg) Then
        Set ws = GetSheet(wb, expectedName)
        If ws Is Nothing Then
            MsgBox "Khong tim thay sheet '" & expectedName & "'.", vbExclamation
            Exit Function
        End If
        ws.Activate
        EnsureActiveSheet = True
    End If
End Function

Public Function ConfirmActiveSheetRisk(Optional ByVal extra As String = "") As Boolean
    Dim msg As String
    If ActiveSheet Is Nothing Then Exit Function
    msg = "Macro nay se thao tac tren sheet hien tai '" & ActiveSheet.Name & "'."
    If extra <> "" Then msg = msg & " " & extra
    msg = msg & " Tiep tuc?"
    ConfirmActiveSheetRisk = (MsgBox(msg, vbYesNo + vbExclamation, "Xac nhan") = vbYes)
End Function
