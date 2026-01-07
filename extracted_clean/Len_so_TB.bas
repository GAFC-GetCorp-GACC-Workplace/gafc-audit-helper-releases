Attribute VB_Name = "Len_so_TB"
Option Explicit

Public Sub Len_so_bao_cao(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub

    Dim wb As Workbook
    Dim wsTB As Worksheet, wsFS As Worksheet
    Dim lastRowTB As Long, lastRowFS As Long
    Dim i As Long, j As Long
    Dim code1 As String, code2 As String
    Dim val1 As Double, val2 As Double
    Dim arrFS As Variant, arrFSVal As Variant
    Dim tbData As Variant
    Dim dictFS As Object
    Dim fsRow As Long
    Dim curVal As Double
    Dim tempVal2 As Double
    Dim fsKey As String

    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim calcMode As XlCalculation

    ' L?y workbook dang active
    Set wb = ActiveSheet.Parent

    ' Ki?m tra sheet
    On Error Resume Next
    Set wsTB = wb.Sheets("TB")
    Set wsFS = wb.Sheets("Adjusted FS")
    On Error GoTo 0

    If wsTB Is Nothing Or wsFS Is Nothing Then
        MsgBox "? Kh�ng t�m th?y sheet 'TB' ho?c 'Adjusted FS'", vbCritical
        Exit Sub
    End If

    ' T?i uu
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    calcMode = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Last row
    lastRowTB = wsTB.Cells(wsTB.Rows.Count, "H").End(xlUp).Row
    lastRowFS = 250   ' c� th? d?i n?u c?n

    ' �?c FS
    arrFS = wsFS.Range("D1:D" & lastRowFS).Value
    arrFSVal = wsFS.Range("G1:G" & lastRowFS).Value

    ' Dictionary map CODE -> ROW
    Set dictFS = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(arrFS, 1)
        fsKey = Trim$(CStr(arrFS(j, 1)))
        If fsKey <> "" Then
            If Not dictFS.Exists(fsKey) Then dictFS.Add fsKey, j
        End If
    Next j

    ' �?c TB
    tbData = wsTB.Range("A2:I" & lastRowTB).Value

    ' ===== X? L� =====
    For i = 1 To UBound(tbData, 1)

        code1 = Trim$(CStr(tbData(i, 1)))
        code2 = Trim$(CStr(tbData(i, 2)))

        val1 = 0: val2 = 0
        If IsNumeric(tbData(i, 8)) Then val1 = tbData(i, 8)
        If IsNumeric(tbData(i, 9)) Then val2 = tbData(i, 9)

        ' Tru?ng h?p d?c bi?t 4211 / 4212
        If (code1 = "4211" Or code1 = "4212") And code1 = code2 Then
            If dictFS.Exists(code1) Then
                fsRow = dictFS(code1)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + (-val1 + val2)
            End If

        Else
            ' CODE1
            If code1 <> "" And dictFS.Exists(code1) Then
                fsRow = dictFS(code1)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + val1
            End If

            ' CODE2 (�m)
            If code2 <> "" And dictFS.Exists(code2) Then
                tempVal2 = val2
                If code2 = "2141" Or code2 = "2142" Or code2 = "2143" _
                   Or code2 = "2417" Or code2 = "139" Or code2 = "159" Then
                    tempVal2 = -val2
                End If

                fsRow = dictFS(code2)
                curVal = 0
                If IsNumeric(arrFSVal(fsRow, 1)) Then curVal = arrFSVal(fsRow, 1)
                arrFSVal(fsRow, 1) = curVal + tempVal2
            End If
        End If
    Next i

    ' ===== GHI K?T QU? � KH�NG �� C�NG TH?C =====
    For j = 1 To lastRowFS

        ' Ch? d�ng c� CODE ? D
        If Trim(wsFS.Cells(j, "D").Value) <> "" Then

            ' Kh�ng d?ng � c� c�ng th?c
            If Not wsFS.Cells(j, "G").HasFormula Then

                ' Ch? ghi khi c� ph�t sinh
                If arrFSVal(j, 1) <> 0 Then
                    wsFS.Cells(j, "G").Value = arrFSVal(j, 1)
                End If

            End If
        End If
    Next j

    ' Tr? tr?ng th�i Excel
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.Calculation = calcMode

    MsgBox "? �� l�n s? Adjusted FS � KH�NG �� C�NG TH?C", vbInformation
End Sub


