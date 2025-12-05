Attribute VB_Name = "Loc_tai_khoan_tao_sheet_TB_2"
Option Explicit
Public Sub Xuly(control As IRibbonControl)
    ' ? Tang t?c
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim wsSrc As Worksheet, wsTB As Worksheet
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim specialDict As Object: Set specialDict = CreateObject("Scripting.Dictionary")
    Dim allTKs As Object: Set allTKs = CreateObject("Scripting.Dictionary")
    Dim prefixSpecialInData As Object: Set prefixSpecialInData = CreateObject("Scripting.Dictionary")
    Set wsSrc = ActiveSheet
    Dim lastRow As Long: lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim tkFull As Variant, tkKey As String
    Dim specialAccounts As Variant, k As Variant, tkSpecial As Variant
    ' ? Danh sách tài kho?n d?c bi?t
    specialAccounts = Array("2141", "2142", "2143", "4211", "4212", "2147", _
                            "2421", "2422", "8211", "8212", "2441", "2442", _
                            "3411", "3412", "2444", "2291", "2292", "2293", "2294")
    ' ? Ðua vào Dictionary
    For i = LBound(specialAccounts) To UBound(specialAccounts)
        specialDict(specialAccounts(i)) = True
    Next i
    ' ? Bu?c 1: Luu toàn b? tài kho?n g?c + dánh d?u prefix d?c bi?t xu?t hi?n
    For i = 2 To lastRow
        tkFull = Trim(wsSrc.Cells(i, 1).Value)
        If tkFull <> "" Then
            If Not allTKs.Exists(tkFull) Then allTKs.Add tkFull, True
            If specialDict.Exists(tkFull) Then
                prefixSpecialInData(Left(tkFull, 3)) = True
            End If
        End If
    Next i
    ' ? Bu?c 2: B? sung mã d?c bi?t n?u có mã con xu?t hi?n
    For Each tkSpecial In specialDict.keys
        If Not allTKs.Exists(tkSpecial) Then
            For Each tkFull In allTKs.keys
                If Left(tkFull, Len(tkSpecial)) = tkSpecial And tkFull <> tkSpecial Then
                    allTKs.Add tkSpecial, True ' thêm mã d?c bi?t
                    prefixSpecialInData(Left(tkSpecial, 3)) = True ' dánh d?u prefix
                    Exit For
                End If
            Next tkFull
        End If
    Next tkSpecial
    ' ? Bu?c 3: L?c tài kho?n cu?i cùng
    For Each tkFull In allTKs.keys
        If specialDict.Exists(tkFull) Then
            tkKey = tkFull ' gi? nguyên d?c bi?t
        Else
            tkKey = Left(tkFull, 3)
            ' N?u dã có mã d?c bi?t cùng prefix ? lo?i mã g?c
            If prefixSpecialInData.Exists(tkKey) Then
                tkKey = ""
            End If
        End If
        If tkKey <> "" Then
            If Not dict.Exists(tkKey) Then dict.Add tkKey, True
        End If
    Next tkFull
    ' ? Xóa sheet TB n?u dã có
    Application.DisplayAlerts = False
    On Error Resume Next: Worksheets("TB").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    ' ? T?o sheet TB m?i
    Set wsTB = Worksheets.Add
    wsTB.Name = "TB"
    ' ? Tiêu d?
    wsTB.Range("A1:I1").Value = Array("Code1", "Code2", "T" & ChrW(224) & "i kho" & ChrW(7843) & "n", "N" & ChrW(7907), "Có", "N" & ChrW(7907), "Có", "N" & ChrW(7907), "Có")
    ' ? Ghi d? li?u
    Dim rowIdx As Long: rowIdx = 2
    For Each k In dict.keys
        wsTB.Cells(rowIdx, 3).Value = k
        rowIdx = rowIdx + 1
    Next k
    ' ? Ð?nh d?ng
    With wsTB.Range("A1:I1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(150, 240, 255)
    End With
    wsTB.Range("A1:I1").AutoFilter
    ' ? B?t l?i
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Done", vbInformation
End Sub

