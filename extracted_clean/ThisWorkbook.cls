Private Sub Workbook_Open()
    If Not GuardLicense() Then
        ' Tam thoi khong dong workbook de debug khi chua kich hoat duoc
        MsgBox "GAFC Audit Helper chua co license hop le. Vui long kich hoat de su dung.", vbExclamation, "License Required"
        Exit Sub
    End If
End Sub
