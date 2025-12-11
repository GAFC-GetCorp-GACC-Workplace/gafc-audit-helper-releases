Attribute VB_Name = "TH_Auto"
Option Explicit
Public gTHEvents As AppEvents_TH
Public Sub Enable_TH_AutoRefresh()
    On Error Resume Next
    Set gTHEvents = New AppEvents_TH
    Set gTHEvents.App = Application
    On Error GoTo 0
End Sub
Public Sub Disable_TH_AutoRefresh()
    On Error Resume Next
    Set gTHEvents = Nothing
    On Error GoTo 0
End Sub
Public Sub Force_TH_Events()
    ' Macro thu cong neu can khoi tao lai Application events
    Enable_TH_AutoRefresh
End Sub
Public Sub TH_Handle_DblClick_Bridge(ByVal Sh As Object, ByVal target As Range, ByRef Cancel As Boolean)
    ' Goi tu Workbook_SheetBeforeDoubleClick de dam bao double-click luon hoat dong
    If gTHEvents Is Nothing Then Enable_TH_AutoRefresh
    If Not gTHEvents Is Nothing Then
        On Error Resume Next
        gTHEvents.Handle_TH_DoubleClick Sh, target, Cancel
        On Error GoTo 0
    End If
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
    ' Dam bao Application events duoc gan truoc khi xu ly
    If gTHEvents Is Nothing Then Enable_TH_AutoRefresh
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
    If ret <> "" Then
        WarnToast ret
    End If
End Sub
