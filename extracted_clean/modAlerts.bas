Attribute VB_Name = "modAlerts"
Option Explicit

' Hiển thị thông báo dạng toast (status bar) và tự tắt sau delaySeconds.
' Dùng cho các thông báo không cần tương tác.
Public Sub Toast(msg As String, Optional title As String = "", Optional delaySeconds As Double = 1#)
    Application.DisplayStatusBar = True
    Application.StatusBar = IIf(title <> "", title & ": " & msg, msg)
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, delaySeconds), "'" & ThisWorkbook.Name & "'!modAlerts.ClearToast"
    On Error GoTo 0
End Sub

Public Sub ClearToast()
    On Error Resume Next
    Application.StatusBar = False
End Sub

' Chuyển các MsgBox không cần nhập liệu sang toast tự tắt.
Public Sub InfoToast(msg As String, Optional delaySeconds As Double = 1#)
    Toast msg, "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", delaySeconds
End Sub

' Thông báo cảnh báo (không bắt click) – dùng khi chỉ nhắc nhẹ.
Public Sub WarnToast(msg As String, Optional delaySeconds As Double = 1#)
    Toast msg, "C" & ChrW(7843) & "nh b" & ChrW(225) & "o", delaySeconds
End Sub
