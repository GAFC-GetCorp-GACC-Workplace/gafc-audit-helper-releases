Attribute VB_Name = "TH_Auto"
Option Explicit
Private gTHEvents As AppEvents_TH
Public Sub Enable_TH_AutoRefresh()
    On Error Resume Next
    Set gTHEvents = New AppEvents_TH
    Set gTHEvents.App = Application
    On Error GoTo 0
End Sub
Public Sub Auto_Open()
    Enable_TH_AutoRefresh
End Sub
Public Sub Workbook_Open()
    Enable_TH_AutoRefresh
End Sub
Public Sub Refresh_TH(Optional wb As Workbook)
    Dim wsNKC As Worksheet
    Dim wsTH As Worksheet
    Dim ret As String
    Dim prevEvents As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsNKC = wb.Worksheets("NKC")
    Set wsTH = wb.Worksheets("TH")
    On Error GoTo 0
    If wsNKC Is Nothing Then Exit Sub
    ' Auto tao TH neu chua co
    If wsTH Is Nothing Then
        On Error Resume Next
        Set wsTH = Tao_TH_Template(wb, wb.Worksheets(wb.Worksheets.Count))
        On Error GoTo 0
    End If
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ret = Auto_Tinh_TH(wsNKC)
    Application.EnableEvents = prevEvents
    ' Only thong bao neu that su co loi
    If ret <> "" Then MsgBox ret, vbExclamation
End Sub
