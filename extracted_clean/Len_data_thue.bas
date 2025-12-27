Attribute VB_Name = "Len_data_thue"
' Bi?n toàn c?c d? d?m file
Option Explicit
Public fileCount As Long
Public maxFiles As Long
' Hàm d? quy t?i uu v?i gi?i h?n và hi?n th? ti?n trình
Private Sub GetAllXMLFiles(ByVal folderPath As String, ByRef fileList As Collection, Optional ByVal depth As Integer = 0)
    Dim fso As Object, folder As Object, subFolder As Object, file As Object
    ' Gi?i h?n d? sâu d? quy d? tránh quá t?i
    Const MAX_DEPTH As Integer = 10
    If depth > MAX_DEPTH Then Exit Sub
    ' Gi?i h?n s? lu?ng file
    If fileList.Count >= maxFiles Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Ki?m tra folder t?n t?i
    If Not fso.FolderExists(folderPath) Then Exit Sub
    On Error Resume Next ' B? qua l?i truy c?p folder
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    ' Hi?n th? ti?n trình
    If fileList.Count Mod 10 = 0 Then
        Application.StatusBar = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(233) & ChrW(116) & ChrW(58) & ChrW(32) & fileList.Count & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(116) & ChrW(236) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7845) & ChrW(121) & ChrW(46) & ChrW(46) & ChrW(46) & ChrW(32) & ChrW(40) & folderPath & ChrW(41)
        DoEvents ' Cho phép Windows x? lý các s? ki?n khác
    End If
    ' L?y file XML trong folder hi?n t?i
    On Error Resume Next
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
            fileList.Add file.path
            If fileList.Count >= maxFiles Then Exit Sub
        End If
    Next
    On Error GoTo 0
    ' Ð? quy vào subfolder
    On Error Resume Next
    For Each subFolder In folder.SubFolders
        ' B? qua các folder h? th?ng
        If Not (subFolder.Name = "System Volume Information" Or _
                subFolder.Name = "$Recycle.Bin" Or _
                Left(subFolder.Name, 1) = ".") Then
            GetAllXMLFiles subFolder.path, fileList, depth + 1
        End If
        If fileList.Count >= maxFiles Then Exit Sub
    Next subFolder
    On Error GoTo 0
End Sub
' Sub chính v?i t?i uu và x? lý l?i
Public Sub tra_cuu_thue(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim wsGTGT As Worksheet, wsTNCN As Worksheet, wsNTNN As Worksheet
    Dim missing As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Set wsGTGT = GetSheet(wb, "GTGT")
    Set wsTNCN = GetSheet(wb, "TNCN")
    Set wsNTNN = GetSheet(wb, "NhaThauNN")
    If wsGTGT Is Nothing And wsTNCN Is Nothing And wsNTNN Is Nothing Then
        MsgBox "Chua co sheet mau thue (GTGT/TNCN/NhaThauNN). Hay chay 'Tra_cuu' de tao mau truoc.", vbExclamation
        Exit Sub
    End If
    If wsGTGT Is Nothing Then missing = missing & "GTGT, "
    If wsTNCN Is Nothing Then missing = missing & "TNCN, "
    If wsNTNN Is Nothing Then missing = missing & "NhaThauNN, "
    If missing <> "" Then
        missing = Left$(missing, Len(missing) - 2)
        If Not ConfirmProceed("Thieu sheet: " & missing & ". Du lieu loai do se bi bo qua. Tiep tuc?") Then Exit Sub
    End If
    Dim xmlDoc As Object
    Dim filePath As Variant, fileList As Collection, pickedFiles As Variant
    Dim folderPath As String
    Dim ws As Worksheet
    Dim r As Long, i As Long
    Dim SelectedMode As String
    Dim HSoKhai As Object, CTieuTKhai As Object
    Dim sheetName As String
    Dim tagPaths_GTGT, tagPaths_TNCN
    Dim rowMap As Object: Set rowMap = CreateObject("Scripting.Dictionary")
    Dim sheetsWithData As Object: Set sheetsWithData = CreateObject("Scripting.Dictionary")
    ' Thi?t l?p gi?i h?n s? file t?i da
    maxFiles = 1000 ' Có th? di?u ch?nh tùy nhu c?u
    ' T?t c?p nh?t màn hình d? tang t?c
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler
    Set fileList = New Collection
    frmSelect.Show
    SelectedMode = frmSelect.SelectedMode
    Unload frmSelect
    If SelectedMode = "" Then GoTo CleanUp
    If SelectedMode = "folder" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = ChrW(67) & ChrW(104) & ChrW(7885) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(432) & ChrW(32) & ChrW(109) & ChrW(7909) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7913) & ChrW(97) & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(32) & ChrW(40) & ChrW(116) & ChrW(7889) & ChrW(105) & ChrW(32) & ChrW(273) & ChrW(97) & ChrW(32) & maxFiles & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(41)
            If .Show = -1 Then
                folderPath = .SelectedItems(1)
                ' Hi?n th? thông báo dang quét
                Application.StatusBar = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(233) & ChrW(116) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(432) & ChrW(32) & ChrW(109) & ChrW(7909) & ChrW(99) & ChrW(44) & ChrW(32) & ChrW(118) & ChrW(117) & ChrW(105) & ChrW(32) & ChrW(108) & ChrW(242) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(7901) & ChrW(46) & ChrW(46) & ChrW(46)
                DoEvents
                ' G?i hàm d? quy d? l?y file XML
                GetAllXMLFiles folderPath, fileList
                If fileList.Count = 0 Then
                    MsgBox ChrW(75) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(116) & ChrW(236) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7845) & ChrW(121) & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(32) & ChrW(110) & ChrW(224) & ChrW(111) & ChrW(33), vbExclamation
                    GoTo CleanUp
                ElseIf fileList.Count >= maxFiles Then
                    If MsgBox(ChrW(272) & ChrW(227) & ChrW(32) & ChrW(116) & ChrW(236) & ChrW(109) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(7845) & ChrW(121) & ChrW(32) & maxFiles & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(32) & ChrW(40) & ChrW(103) & ChrW(105) & ChrW(7899) & ChrW(105) & ChrW(32) & ChrW(104) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(7889) & ChrW(105) & ChrW(32) & ChrW(273) & ChrW(97) & ChrW(41) & ChrW(46) & vbCrLf & _
                             ChrW(66) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(109) & ChrW(117) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7909) & ChrW(99) & ChrW(32) & ChrW(120) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(253) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(63), vbYesNo + vbQuestion) = vbNo Then
                        GoTo CleanUp
                    End If
                End If
            Else
                GoTo CleanUp
            End If
        End With
    ElseIf SelectedMode = "file" Then
        pickedFiles = Application.GetOpenFilename("XML Files (*.xml), *.xml", MultiSelect:=True)
        If VarType(pickedFiles) = vbBoolean Then GoTo CleanUp
        If IsArray(pickedFiles) Then
            For i = LBound(pickedFiles) To UBound(pickedFiles)
                fileList.Add pickedFiles(i)
            Next i
        Else
            fileList.Add pickedFiles
        End If
    End If
    ' Kh?i t?o arrays
    tagPaths_GTGT = Array( _
        "ns:ct22", _
        "ns:GiaTriVaThueGTGTHHDVMuaVao/ns:ct23", _
        "ns:HangHoaDichVuNhapKhau/ns:ct23a", _
        "ns:GiaTriVaThueGTGTHHDVMuaVao/ns:ct24", _
        "ns:HangHoaDichVuNhapKhau/ns:ct24a", _
        "ns:ct25", "ns:ct26", "ns:ct29", _
        "ns:HHDVBRaChiuTSuat5/ns:ct30", _
        "ns:HHDVBRaChiuTSuat10/ns:ct32", "ns:ct32a", _
        "ns:TongDThuVaThueGTGTHHDVBRa/ns:ct34", _
        "ns:HHDVBRaChiuTSuat5/ns:ct31", _
        "ns:HHDVBRaChiuTSuat10/ns:ct33", _
        "ns:TongDThuVaThueGTGTHHDVBRa/ns:ct35", _
        "ns:ct36", "ns:ct37", "ns:ct38", _
        "ns:ct40", "ns:ct41", "ns:ct42", "ns:ct43")
    tagPaths_TNCN = Array( _
        "ns:CTieuTKhaiChinh/ns:ct16", "ns:CTieuTKhaiChinh/ns:ct17", "ns:CTieuTKhaiChinh/ns:ct19", _
        "ns:CTieuTKhaiChinh/ns:ct20", "ns:CTieuTKhaiChinh/ns:ct18", "ns:CTieuTKhaiChinh/ns:ct22", _
        "ns:CTieuTKhaiChinh/ns:ct23", "ns:CTieuTKhaiChinh/ns:ct21", "ns:CTieuTKhaiChinh/ns:ct27", _
        "ns:CTieuTKhaiChinh/ns:ct28", "ns:CTieuTKhaiChinh/ns:ct26", "ns:CTieuTKhaiChinh/ns:ct30", _
        "ns:CTieuTKhaiChinh/ns:ct31", "ns:CTieuTKhaiChinh/ns:ct29")
    ' X? lý t?ng file v?i progress
    Dim fileCounter As Long
    fileCounter = 0
    For Each filePath In fileList
        fileCounter = fileCounter + 1
        ' C?p nh?t ti?n trình m?i 5 file
        If fileCounter Mod 5 = 0 Or fileCounter = fileList.Count Then
            Application.StatusBar = ChrW(272) & ChrW(97) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(120) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(253) & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & fileCounter & ChrW(47) & fileList.Count & ChrW(46) & ChrW(46) & ChrW(46)
            DoEvents
        End If
        Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
        xmlDoc.Async = False
        xmlDoc.SetProperty "SelectionNamespaces", "xmlns:ns='http://kekhaithue.gdt.gov.vn/TKhaiThue'"
        If Not xmlDoc.Load(CStr(filePath)) Then GoTo NextFile
        Set HSoKhai = xmlDoc.SelectSingleNode("//ns:HSoKhaiThue")
        If HSoKhai Is Nothing Then GoTo NextFile
        Set CTieuTKhai = HSoKhai.SelectSingleNode("ns:CTieuTKhaiChinh")
        Dim tenNNT As String, kyKKhai As String, kieuKy As String, soLanKhai As String, tenTKhai As String
        tenNNT = GetNodeText(HSoKhai, "ns:TTinChung/ns:TTinTKhaiThue/ns:NNT/ns:tenNNT")
        kyKKhai = GetNodeText(HSoKhai, "ns:TTinChung/ns:TTinTKhaiThue/ns:TKhaiThue/ns:KyKKhaiThue/ns:kyKKhai")
        kieuKy = GetNodeText(HSoKhai, "ns:TTinChung/ns:TTinTKhaiThue/ns:TKhaiThue/ns:KyKKhaiThue/ns:kieuKy")
        soLanKhai = GetNodeText(HSoKhai, "ns:TTinChung/ns:TTinTKhaiThue/ns:TKhaiThue/ns:soLan")
        tenTKhai = GetNodeText(HSoKhai, "ns:TTinChung/ns:TTinTKhaiThue/ns:TKhaiThue/ns:tenTKhai")
        If tenTKhai Like "*GTGT*" Then
            sheetName = "GTGT"
        ElseIf tenTKhai Like "*TNCN*" Then
            sheetName = "TNCN"
        ElseIf tenTKhai Like "*NTNN*" Then
            sheetName = "NhaThauNN"
        Else
            GoTo NextFile
        End If
        On Error Resume Next
        Set ws = ActiveWorkbook.Sheets(sheetName)
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextFile
        r = 5
        Do While Application.WorksheetFunction.CountA(ws.Range("A" & r & ":Y" & r)) > 0
            r = r + 1
        Loop
        ' Ðánh d?u sheet có data
        If Not sheetsWithData.Exists(sheetName) Then
            sheetsWithData.Add sheetName, True
        End If
        With ws
            .Cells(r, 1).Value = tenNNT
            If kieuKy = "Q" Then
                .Cells(r, 2).Value = ChrW(81) & ChrW(117) & ChrW(253) & ChrW(32) & kyKKhai
            ElseIf kieuKy = "M" Then
                .Cells(r, 2).Value = ChrW(84) & ChrW(104) & ChrW(225) & ChrW(110) & ChrW(103) & ChrW(32) & kyKKhai
            ElseIf kieuKy = "D" Then
                .Cells(r, 2).Value = ChrW(78) & ChrW(103) & ChrW(224) & ChrW(121) & ChrW(32) & kyKKhai
            Else
                .Cells(r, 2).Value = kyKKhai
            End If
            If soLanKhai = "0" Then
                .Cells(r, 3).Value = ChrW(76) & ChrW(7847) & ChrW(110) & ChrW(32) & ChrW(273) & ChrW(7847) & ChrW(117) & ChrW(32) & ChrW(42)
            ElseIf soLanKhai <> "" Then
                .Cells(r, 3).Value = ChrW(66) & ChrW(7893) & ChrW(32) & ChrW(115) & ChrW(117) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(108) & ChrW(7847) & ChrW(110) & ChrW(32) & soLanKhai & ChrW(32) & ChrW(42)
            End If
            Select Case sheetName
                Case "GTGT"
                    For i = 0 To UBound(tagPaths_GTGT)
                        .Cells(r, i + 4).Value = GetNodeText(CTieuTKhai, tagPaths_GTGT(i))
                    Next i
                    .Range("D" & r & ":Y" & r).NumberFormat = "#,##0;(#,##0)"
                Case "TNCN"
                    For i = 0 To UBound(tagPaths_TNCN)
                        .Cells(r, i + 4).Value = GetNodeText(HSoKhai, tagPaths_TNCN(i))
                    Next i
                    .Range("D" & r & ":Q" & r).NumberFormat = "#,##0;(#,##0)"
                Case "NhaThauNN"
                    Dim startRow As Long
                    startRow = r
                    .Cells(r, 5).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct7")
                    .Cells(r, 7).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct9")
                    .Cells(r, 8).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct10")
                    .Cells(r, 10).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct12")
                    .Cells(r, 11).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct13")
                    .Cells(r, 12).Value = GetNodeText(HSoKhai, "ns:CTieuTKhaiChinh/ns:tong_ct14")
                    .Range("A" & r & ":W" & r).Font.Bold = True
                    Dim ThueNodes As Object, ThueNode As Object
                    Dim rChiTiet As Long
                    Set ThueNodes = HSoKhai.SelectNodes("ns:CTieuTKhaiChinh/ns:BKThueNTNN/ns:ThueNTNN")
                    If Not ThueNodes Is Nothing Then
                        For i = 0 To ThueNodes.Length - 1
                            Set ThueNode = ThueNodes.item(i)
                            r = r + 1
                            rChiTiet = r
                            .Cells(r, 1).Value = "   " & GetNodeText(ThueNode, "ns:ct1b")
                            .Cells(r, 4).Value = GetNodeText(ThueNode, "ns:ct5")
                            .Cells(r, 5).Value = GetNodeText(ThueNode, "ns:ThueGTGT/ns:ct7")
                            .Cells(r, 6).Value = GetNodeText(ThueNode, "ns:ThueGTGT/ns:ct8")
                            .Cells(r, 7).Value = GetNodeText(ThueNode, "ns:ThueGTGT/ns:ct9")
                            .Cells(r, 8).Value = GetNodeText(ThueNode, "ns:ThueTNDN/ns:ct10")
                            .Cells(r, 9).Value = GetNodeText(ThueNode, "ns:ThueTNDN/ns:ct11")
                            .Cells(r, 10).Value = GetNodeText(ThueNode, "ns:ThueTNDN/ns:ct12")
                            .Cells(r, 11).Value = GetNodeText(ThueNode, "ns:ThueTNDN/ns:ct13")
                            .Cells(r, 12).Value = GetNodeText(ThueNode, "ns:ct14")
                        Next i
                        With .Range("A" & rChiTiet & ":L" & rChiTiet).Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .Weight = xlThin
                            .Color = RGB(150, 150, 150)
                        End With
                    End If
                    .Range("D" & startRow & ":L" & r).NumberFormat = "#,##0;(#,##0)"
            End Select
            If soLanKhai <> "" And soLanKhai <> "0" Then
                Dim key As String
                key = sheetName & "|" & tenNNT & "|" & kyKKhai & "|" & kieuKy
                If Not rowMap.Exists(key) Then
                    rowMap.Add key, Array(r, CLng(soLanKhai))
                ElseIf CLng(soLanKhai) > rowMap(key)(1) Then
                    rowMap(key) = Array(r, CLng(soLanKhai))
                End If
            End If
        End With
NextFile:
        ' Gi?i phóng b? nh?
        Set xmlDoc = Nothing
        Set HSoKhai = Nothing
        Set CTieuTKhai = Nothing
    Next filePath
    ' Highlight và d?i màu tab
    Dim entry
    For Each entry In rowMap
        Dim arr: arr = Split(entry, "|")
        ActiveWorkbook.Sheets(arr(0)).Range("A" & rowMap(entry)(0) & ":Y" & rowMap(entry)(0)).Interior.Color = RGB(255, 255, 153)
    Next
    Dim sheetKey As Variant
    For Each sheetKey In sheetsWithData.keys
        On Error Resume Next
        ActiveWorkbook.Sheets(sheetKey).Tab.Color = RGB(255, 0, 0)
        On Error GoTo 0
    Next sheetKey
    ' Thông báo k?t qu?
    Dim importedSheets As String
    importedSheets = ""
    If sheetsWithData.Count > 0 Then
        For Each sheetKey In sheetsWithData.keys
            If importedSheets = "" Then
                importedSheets = sheetKey
            Else
                importedSheets = importedSheets & ", " & sheetKey
            End If
        Next sheetKey
        MsgBox ChrW(272) & ChrW(227) & ChrW(32) & ChrW(105) & ChrW(109) & ChrW(112) & ChrW(111) & ChrW(114) & ChrW(116) & ChrW(32) & ChrW(120) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & fileList.Count & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(46) & vbCrLf & _
               ChrW(67) & ChrW(225) & ChrW(99) & ChrW(32) & ChrW(115) & ChrW(104) & ChrW(101) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(105) & ChrW(109) & ChrW(112) & ChrW(111) & ChrW(114) & ChrW(116) & ChrW(58) & ChrW(32) & importedSheets & vbCrLf & _
               ChrW(40) & ChrW(84) & ChrW(97) & ChrW(98) & ChrW(32) & ChrW(109) & ChrW(224) & ChrW(117) & ChrW(32) & ChrW(273) & ChrW(7887) & ChrW(32) & ChrW(108) & ChrW(224) & ChrW(32) & ChrW(115) & ChrW(104) & ChrW(101) & ChrW(101) & ChrW(116) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(109) & ChrW(7899) & ChrW(105) & ChrW(41), vbInformation
    Else
        MsgBox ChrW(272) & ChrW(227) & ChrW(32) & ChrW(113) & ChrW(117) & ChrW(233) & ChrW(116) & ChrW(32) & fileList.Count & ChrW(32) & ChrW(102) & ChrW(105) & ChrW(108) & ChrW(101) & ChrW(32) & ChrW(88) & ChrW(77) & ChrW(76) & ChrW(32) & ChrW(110) & ChrW(104) & ChrW(432) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(107) & ChrW(104) & ChrW(244) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(99) & ChrW(243) & ChrW(32) & ChrW(100) & ChrW(97) & ChrW(116) & ChrW(97) & ChrW(32) & ChrW(112) & ChrW(104) & ChrW(249) & ChrW(32) & ChrW(104) & ChrW(7907) & ChrW(112) & ChrW(46), vbInformation
    End If
CleanUp:
    ' Khôi ph?c l?i thi?t l?p
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ' Gi?i phóng b? nh?
    Set fileList = Nothing
    Set rowMap = Nothing
    Set sheetsWithData = Nothing
    Exit Sub
ErrorHandler:
    MsgBox ChrW(76) & ChrW(7895) & ChrW(105) & ChrW(58) & ChrW(32) & Err.Description, vbCritical
    Resume CleanUp
End Sub
' Function GetNodeText - gi? nguyên
Function GetNodeText(xmlNode As Object, ByVal path As String) As String
    On Error Resume Next
    Dim resultNode As Object
    Set resultNode = xmlNode.SelectSingleNode(path)
    If Not resultNode Is Nothing Then
        GetNodeText = resultNode.text
    Else
        GetNodeText = ""
    End If
    On Error GoTo 0
End Function
' Sub Xoa_data v?i Unicode
Public Sub Xoa_data(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    If Not ConfirmActiveSheetRisk("Du lieu se bi xoa tu dong 5 den cuoi.") Then Exit Sub
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then lastRow = 5
    With ws.Range("A5:Y" & lastRow)
        .UnMerge
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    On Error Resume Next
    ws.Tab.ColorIndex = xlColorIndexNone
    On Error GoTo 0
    MsgBox ChrW(272) & ChrW(227) & ChrW(32) & ChrW(120) & ChrW(243) & ChrW(97) & ChrW(33), vbInformation
End Sub

