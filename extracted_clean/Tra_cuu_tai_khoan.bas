Option Explicit
Public Sub Tracode(control As IRibbonControl)
    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim inputArr As Variant, outputArr As Variant
    Dim tempVal As Variant
    Dim tk As String
    Set ws = Worksheets("TB")
    Set dict = CreateObject("Scripting.Dictionary")
    ' ? B?ng tra mã
    dict.Add "111", Array("111", "")
    dict.Add "112", Array("112", "")
    dict.Add "113", Array("113", "")
    dict.Add "121", Array("1211", "")
    dict.Add "128", Array("1281", "")
    dict.Add "131", Array("131", "332")
    dict.Add "133", Array("1331", "")
    dict.Add "136", Array("136", "")
    dict.Add "138", Array("138", "")
    dict.Add "141", Array("141", "")
    dict.Add "151", Array("151", "")
    dict.Add "152", Array("152", "")
    dict.Add "153", Array("153", "")
    dict.Add "154", Array("154", "")
    dict.Add "155", Array("155", "")
    dict.Add "156", Array("156", "")
    dict.Add "157", Array("157", "")
    dict.Add "158", Array("158", "")
    dict.Add "171", Array("171", "")
    dict.Add "211", Array("211", "")
    dict.Add "212", Array("212", "")
    dict.Add "213", Array("213", "")
    dict.Add "2141", Array("", "2141")
    dict.Add "2142", Array("", "2142")
    dict.Add "2143", Array("", "2143")
    dict.Add "217", Array("217", "")
    dict.Add "2147", Array("", "2417")
    dict.Add "221", Array("221", "")
    dict.Add "222", Array("222", "")
    dict.Add "228", Array("228", "")
    dict.Add "2291", Array("", "2291")
    dict.Add "2292", Array("", "229")
    dict.Add "2293", Array("", "139")
    dict.Add "2294", Array("", "159")
    dict.Add "241", Array("241", "")
    dict.Add "242", Array("2421", "")
    dict.Add "2421", Array("2421", "")
    dict.Add "2422", Array("2422", "")
    dict.Add "243", Array("243", "")
    dict.Add "244", Array("2441", "")
    dict.Add "2441", Array("2441", "")
    dict.Add "2442", Array("2442", "")
    dict.Add "2444", Array("2441", "")
    dict.Add "331", Array("132", "331")
    dict.Add "333", Array("1338", "333")
    dict.Add "334", Array("138", "334")
    dict.Add "335", Array("", "335")
    dict.Add "336", Array("136", "336")
    dict.Add "337", Array("", "337")
    dict.Add "338", Array("138", "3388")
    dict.Add "341", Array("", "3411")
    dict.Add "3411", Array("", "3411")
    dict.Add "3412", Array("", "3412")
    dict.Add "343", Array("", "3432")
    dict.Add "344", Array("", "3388")
    dict.Add "347", Array("", "347")
    dict.Add "352", Array("", "352")
    dict.Add "353", Array("", "353")
    dict.Add "356", Array("", "356")
    dict.Add "357", Array("", "357")
    dict.Add "411", Array("", "4111c")
    dict.Add "412", Array("", "412")
    dict.Add "413", Array("", "413")
    dict.Add "414", Array("414", "")
    dict.Add "417", Array("", "417")
    dict.Add "418", Array("", "418")
    dict.Add "4211", Array("4211", "4211")
    dict.Add "4212", Array("4212", "4212")
    dict.Add "511", Array("", "511")
    dict.Add "515", Array("", "515")
    dict.Add "521", Array("", "531")
    dict.Add "632", Array("632", "")
    dict.Add "635", Array("635", "")
    dict.Add "641", Array("641", "")
    dict.Add "642", Array("642", "")
    dict.Add "711", Array("711", "")
    dict.Add "811", Array("811", "")
    dict.Add "821", Array("8211", "")
    dict.Add "8211", Array("8211", "")
    dict.Add "8212", Array("8212", "")
    ' ? Tang t?c VBA
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' ?? Tìm s? dòng
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ' ?? Ð?c toàn b? d? li?u c?t C vào m?ng (tài kho?n)
inputArr = ws.Range("C2:C" & lastRow).Value
    ReDim outputArr(1 To UBound(inputArr), 1 To 2) ' Code1, Code2
    ' ?? X? lý t?ng dòng trong m?ng
    For i = 1 To UBound(inputArr)
        tk = Trim(CStr(inputArr(i, 1))) ' Ép v? chu?i
        If dict.Exists(tk) Then
            tempVal = dict(tk)
            outputArr(i, 1) = tempVal(0)
            outputArr(i, 2) = tempVal(1)
        Else
            outputArr(i, 1) = ""
            outputArr(i, 2) = ""
        End If
    Next i
    ' ?? Ghi toàn b? k?t qu? ra c?t A:B m?t l?n
    ws.Range("A2:B" & lastRow).Value = outputArr
    ' ?? Khôi ph?c cài d?t
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "?? Ðã tra c?u xong Code1 và Code2 v?i t?c d? t?i uu!", vbInformation
End Sub