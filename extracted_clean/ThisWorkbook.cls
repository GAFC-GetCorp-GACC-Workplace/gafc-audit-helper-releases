VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Gan Application events de bat double-click, auto refresh TH, etc.
    On Error Resume Next
    Enable_TH_AutoRefresh
    On Error GoTo 0

    ' Chỉ set trạng thái license, không chặn Excel hoặc bật MsgBox
    If Not InitLicenseStateOnOpen(True) Then
        Exit Sub
    End If

    ' Check for updates (silently, in background)
    On Error Resume Next
    CheckForUpdates
    On Error GoTo 0
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Disable_TH_AutoRefresh
    ApplyPendingUpdateIfNeeded
    On Error GoTo 0
End Sub
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Gọi trực tiếp handler để tránh phụ thuộc vào WithEvents nếu bị rớt
    If Sh Is Nothing Then Exit Sub
    If Sh.Name <> "TH" Then Exit Sub
    Dim handler As Object
    On Error Resume Next
    Set handler = New AppEvents_TH
    If Not handler Is Nothing Then handler.Handle_TH_DoubleClick Sh, Target, Cancel
    On Error GoTo 0
End Sub
