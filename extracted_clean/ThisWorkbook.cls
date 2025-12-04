Private Sub Workbook_Open()
    ' Check license first
    If Not GuardLicense() Then
        ' Tam thoi khong dong workbook de debug khi chua kich hoat duoc
        MsgBox "GAFC Audit Helper chua co license hop le. Vui long kich hoat de su dung.", vbExclamation, "License Required"
        Exit Sub
    End If

    ' Check for updates (silently, in background)
    On Error Resume Next
    CheckForUpdates
    On Error GoTo 0
End Sub
