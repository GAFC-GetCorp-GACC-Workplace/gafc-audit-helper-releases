Attribute VB_Name = "modRibbon"
Option Explicit

' Ribbon callback - Kiểm tra cập nhật
Public Sub OnCheckUpdate(control As IRibbonControl)
    On Error Resume Next
    InfoToast "Checking for updates..."
    CheckForUpdates True  ' Force check ngay
End Sub

' Ribbon callback - Giới thiệu
Public Sub OnShowAbout(control As IRibbonControl)
    Dim msg As String
    Dim currentVer As String

    On Error Resume Next
    ' Get version from GetCurrentVersion() in modAutoUpdate
    currentVer = GetCurrentVersion()

    ' Final fallback
    If Err.Number <> 0 Or Len(currentVer) = 0 Then
        currentVer = "1.0.17"
    End If
    On Error GoTo 0

    msg = "GAFC Audit Helper" & vbCrLf & _
          "Version: " & currentVer & vbCrLf & vbCrLf & _
          "Developed by" & vbCrLf & _
          "GLOBAL AUDITING AND FINANCIAL" & vbCrLf & _
          "CONSULTANCY COMPANY LIMITED" & vbCrLf & vbCrLf & _
          "Contact:" & vbCrLf & _
          "Email: info@globalauditing.com" & vbCrLf & _
          "Hotline: 0918 70 85 72" & vbCrLf & vbCrLf & _
          "(c) 2025 GAFC"

    MsgBox msg, vbInformation, "About"
End Sub
