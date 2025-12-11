Attribute VB_Name = "Tao_format_TK_thue"
Option Explicit
Public Sub Tra_cuu(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
    Dim wb As Workbook
    Dim ws As Worksheet
    '--- T?o workbook m?i ---
    Set wb = Workbooks.Add
    '--- G?i 3 sheet ---
    Call TaoBangThueGTGT_Merge_InWorkbook(wb)
    Call TaoBangTNCN_Merge_InWorkbook(wb)
    Call TaoBangThueNTNN_Merge_InWorkbook(wb)
    '--- Xoá sheet m?c d?nh (Sheet1, Sheet2...) ---
    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name Like "Sheet*" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
    InfoToast "Done"
End Sub
'=====================================================
' Các sub con gi? nguyên code b?n có, ch? thêm d?i s? wb
'=====================================================
Sub TaoBangThueGTGT_Merge_InWorkbook(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add
    ws.Name = "GTGT"
    ws.Range("A2:A3").Merge
    ws.Range("B2:B3").Merge
    ws.Range("C2:C3").Merge
    ws.Range("D2:D3").Merge
    ws.Range("E2:F2").Merge
    ws.Range("G2:H2").Merge
    ws.Range("I2:I3").Merge
    ws.Range("J2:O2").Merge
    ws.Range("P2:R2").Merge
    ws.Range("S2:S3").Merge
    ws.Range("T2:U2").Merge
    ws.Range("V2:V3").Merge
    ws.Range("W2:W3").Merge
    ws.Range("X2:X3").Merge
    ws.Range("Y2:Y3").Merge
    ws.Range("A1:Y1").Merge
    '--- Nh?p d? li?u sau khi g?p ---
    ws.Range("A1").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P S" & ChrW(7888) & " LI" & ChrW(7878) & "U K" & ChrW(202) & " KHAI THU" & ChrW(7870) & " GI" & ChrW(193) & " TR" & ChrW(7882) & " GIA T" & ChrW(258) & "NG"
    ws.Range("A2").Value = "T" & ChrW(234) & "n c" & ChrW(244) & "ng ty"
    ws.Range("B2").Value = "K" & ChrW(7923) & " t" & ChrW(237) & "nh thu" & ChrW(7871)
    ws.Range("C2").Value = "L" & ChrW(7847) & "n k" & ChrW(234) & " khai"
    ws.Range("D2").Value = "K" & ChrW(7923) & " tr" & ChrW(432) & ChrW(7899) & "c chuy" & ChrW(7875) & "n sang"
    ws.Range("E2").Value = "Gi" & ChrW(225) & " tr" & ChrW(7883) & " HH mua v" & ChrW(224) & "o"
    ws.Range("G2").Value = "Thu" & ChrW(7871) & " GTGT " & ChrW(273) & ChrW(7847) & "u v" & ChrW(224) & "o"
    ws.Range("I2").Value = ChrW(272) & ChrW(432) & ChrW(7907) & "c kh" & ChrW(7845) & "u tr" & ChrW(7915)
    ws.Range("J2").Value = "Doanh thu"
    ws.Range("P2").Value = "Thu" & ChrW(7871) & " GTGT"
    ws.Range("S2").Value = "Thu" & ChrW(7871) & " ph" & ChrW(225) & "t sinh trong k" & ChrW(236)
    ws.Range("T2").Value = ChrW(272) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh"
    ws.Range("V2").Value = "C" & ChrW(242) & "n ph" & ChrW(7843) & "i n" & ChrW(7897) & "p"
    ws.Range("W2").Value = "Ch" & ChrW(432) & "a kh" & ChrW(7845) & "u tr" & ChrW(249) & " h" & ChrW(7871) & "t k" & ChrW(7923) & " n" & ChrW(224) & "y"
    ws.Range("X2").Value = ChrW(272) & ChrW(7873) & " ngh" & ChrW(7883) & " ho" & ChrW(224) & "n"
    ws.Range("Y2").Value = "Chuy" & ChrW(7875) & "n k" & ChrW(236) & " sau"
    '--- Sub-headers ---
    ws.Range("E3").Value = "Gi" & ChrW(225) & " tr" & ChrW(7883) & " v" & ChrW(224) & " thu" & ChrW(7871) & " GTGT c" & ChrW(7911) & "a HHDV mua v" & ChrW(224) & "o"
    ws.Range("F3").Value = "HHDV nh" & ChrW(7853) & "p kh" & ChrW(7849) & "u"
    ws.Range("G3").Value = "Gi" & ChrW(225) & " tr" & ChrW(7883) & " v" & ChrW(224) & " thu" & ChrW(7871) & " GTGT c" & ChrW(7911) & "a HHDV mua v" & ChrW(224) & "o"
    ws.Range("H3").Value = "HHDV nh" & ChrW(7853) & "p kh" & ChrW(7849) & "u"
    ws.Range("J3").Value = "Kh" & ChrW(244) & "ng ch" & ChrW(7883) & "u thu" & ChrW(7871)
    ws.Range("K3").Value = "Thu" & ChrW(7871) & " su" & ChrW(7845) & "t 0%"
    ws.Range("L3").Value = "Thu" & ChrW(7871) & " su" & ChrW(7845) & "t 5%"
    ws.Range("M3").Value = "Thu" & ChrW(7871) & " su" & ChrW(7845) & "t 10%"
    ws.Range("N3").Value = "Kh" & ChrW(244) & "ng ch" & ChrW(7883) & "u thu" & ChrW(7871)
    ws.Range("O3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ws.Range("P3").Value = "Thu" & ChrW(7871) & " su" & ChrW(7845) & "t 5%"
    ws.Range("Q3").Value = "Thu" & ChrW(7871) & " su" & ChrW(7845) & "t 10%"
    ws.Range("R3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ws.Range("T3").Value = ChrW(272) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh gi" & ChrW(7843) & "m"
    ws.Range("U3").Value = ChrW(272) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh t" & ChrW(259) & "ng"
    '--- Hàng 4: ctXX ---
    Dim arr As Variant
    arr = Array("/ct22", "/ct23", "/ct23a", "/ct24", "/ct24a", "/ct25", "/ct26", "/ct29", "/ct30", "/ct32", "/ct32a", "/ct34", "/ct31", "/ct33", "/ct35", "/ct36", "/ct37", "/ct38", "/ct40", "/ct41", "/ct42", "/ct43")
    ws.Range("D4:Y4").Value = arr
    '--- Format ---
    ws.Range("A1:Y4").HorizontalAlignment = xlCenter
    ws.Range("A1:Y4").VerticalAlignment = xlCenter
    ws.Range("A1:Y4").WrapText = True
    ws.Range("A1:Y4").Font.Bold = True
    ws.Range("A1:Y4").Borders.LineStyle = xlContinuous
    ws.Columns.AutoFit
    Rows("4").Hidden = True
End Sub
Sub TaoBangTNCN_Merge_InWorkbook(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add
    ws.Name = "TNCN"
    '--- Code c?a b?n gi? nguyên t? dây ---
    ws.Range("A2:A3").Merge
    ws.Range("B2:B3").Merge
    ws.Range("C2:C3").Merge
    ws.Range("D2:D3").Merge
    ws.Range("E2:E3").Merge
    ws.Range("F2:H2").Merge
    ws.Range("I2:K2").Merge
    ws.Range("L2:N2").Merge
    ws.Range("O2:Q2").Merge
    ws.Range("A1:Q1").Merge
    '--- Nh?p d? li?u sau khi merge ---
    ws.Range("A1").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P S" & ChrW(7888) & " LI" & ChrW(7878) & "U K" & ChrW(202) & " KHAI THU" & ChrW(7870) & " THU NH" & ChrW(7852) & "P C" & ChrW(193) & " NH" & ChrW(194) & "N T" & ChrW(7914) & " TI" & ChrW(7872) & "N L" & ChrW(431) & ChrW(416) & "NG, TI" & ChrW(7872) & "N C" & ChrW(212) & "NG"
    ws.Range("A2").Value = "Tên công ty"
    ws.Range("B2").Value = "K" & ChrW(7923) & " t" & ChrW(237) & "nh thu" & ChrW(7871)
    ws.Range("C2").Value = "L" & ChrW(7847) & "n k" & ChrW(234) & " khai"
    ws.Range("D2").Value = "T" & ChrW(7893) & "ng s" & ChrW(7889) & " lao " & ChrW(273) & ChrW(7897) & "ng"
    ws.Range("E2").Value = "Lao " & ChrW(273) & ChrW(7897) & "ng c" & ChrW(432) & " tr" & ChrW(250) & " c" & ChrW(243) & " H" & ChrW(272)
    ws.Range("F2").Value = "S" & ChrW(7889) & " c" & ChrW(225) & " nh" & ChrW(226) & "n kh" & ChrW(7845) & "u tr" & ChrW(7915) & " thu" & ChrW(7871)
    ws.Range("I2").Value = "Thu nh" & ChrW(7853) & "p ch" & ChrW(7883) & "u thu" & ChrW(7871) & " " & ChrW(273) & ChrW(227) & " tr" & ChrW(7843)
    ws.Range("L2").Value = "Thu nh" & ChrW(7853) & "p ch" & ChrW(7883) & "u thu" & ChrW(7871) & " " & ChrW(273) & ChrW(227) & " tr" & ChrW(7843) & " cho c" & ChrW(225) & " nh" & ChrW(226) & "n thu" & ChrW(7897) & "c di" & ChrW(7879) & "n kh" & ChrW(7845) & "u tr" & ChrW(7915) & " thu" & ChrW(7871)
    ws.Range("O2").Value = "Thu" & ChrW(7871) & " TNCN " & ChrW(273) & ChrW(227) & " kh" & ChrW(7845) & "u tr" & ChrW(7915)
    ws.Range("O2").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P S" & ChrW(7888) & " LI" & ChrW(7878) & "U K" & ChrW(202) & " KHAI THU" & ChrW(7870) & " THU NH" & ChrW(7852) & "P C" & ChrW(193) & " NH" & ChrW(194) & "N T" & ChrW(7914) & " TI" & ChrW(7872) & "N L" & ChrW(431) & ChrW(416) & "NG, TI" & ChrW(7872) & "N C" & ChrW(212) & "NG"
    '--- Hàng 3: Sub-headers ---
    ws.Range("F3").Value = "C" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("G3").Value = "Kh" & ChrW(244) & "ng c" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("H3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ws.Range("I3").Value = "C" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("J3").Value = "Kh" & ChrW(244) & "ng c" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("K3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ws.Range("L3").Value = "C" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("M3").Value = "Kh" & ChrW(244) & "ng c" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("N3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ws.Range("O3").Value = "C" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("P3").Value = "Kh" & ChrW(244) & "ng c" & ChrW(432) & " tr" & ChrW(250)
    ws.Range("Q3").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    '--- Hàng 4: Ch? s? ---
    Dim arr As Variant
    arr = Array( _
        "394:/ct21|864:/ct16", _
        "394:/ct22|864:/ct17", _
        "394:/ct24|864:/ct19", _
        "394:/ct25|864:/ct20", _
        "394:/ct23|864:/ct18", _
        "394:/ct27|864:/ct22", _
        "394:/ct28|864:/ct23", _
        "394:/ct26|864:/ct21", _
        "394:/ct30|864:/ct27", _
        "394:/ct31|864:/ct28", _
        "394:/ct29|864:/ct26", _
        "394:/ct33|864:/ct30", _
        "394:/ct34|864:/ct31", _
        "394:/ct32|864:/ct29")
    ws.Range("D4:Q4").Value = arr
    '--- Ð?nh d?ng ---
    With ws.Range("A1:Q4")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Columns("A:Q").AutoFit
    Rows("4").Hidden = True
End Sub
Sub TaoBangThueNTNN_Merge_InWorkbook(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add
    ws.Name = "NhaThauNN"
    ws.Range("A2:A3").Merge
    ws.Range("B2:B3").Merge
    ws.Range("C2:C3").Merge
    ws.Range("D2:D3").Merge
    ws.Range("E2:G2").Merge
    ws.Range("H2:K2").Merge
    ws.Range("L2:L3").Merge
    ws.Range("A1:L1").Merge
    '--- Nh?p d? li?u sau khi merge ---
    ws.Range("A2").Value = "Tên công ty"
    ws.Range("B2").Value = "K" & ChrW(7923) & " t" & ChrW(237) & "nh thu" & ChrW(7871)
    ws.Range("C2").Value = "L" & ChrW(7847) & "n k" & ChrW(234) & " khai"
    ws.Range("D2").Value = "Doanh thu ch" & ChrW(432) & "a bao g" & ChrW(7891) & "m thu" & ChrW(7871) & " GTGT"
    ws.Range("E2").Value = "Thu" & ChrW(7871) & " GTGT"
    ws.Range("H2").Value = "Thu" & ChrW(7871) & " TNDN"
    ws.Range("L2").Value = "T" & ChrW(7893) & "ng s" & ChrW(7889) & " thu" & ChrW(7871) & " ph" & ChrW(7843) & "i n" & ChrW(7897) & "p"
    ws.Range("A1").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P S" & ChrW(7888) & " LI" & ChrW(7878) & "U K" & ChrW(202) & " KHAI THU" & ChrW(7870) & " NH" & ChrW(192) & " TH" & ChrW(7846) & "U N" & ChrW(431) & ChrW(7898) & "C NGO" & ChrW(192) & "I"
    '--- Hàng 3: Sub-headers ---
    ws.Range("E3").Value = "Doanh thu t" & ChrW(237) & "nh thu" & ChrW(7871)
    ws.Range("F3").Value = "T" & ChrW(7927) & " l" & ChrW(7879) & " GTGT (%)"
    ws.Range("G3").Value = "Thu" & ChrW(7871) & " GTGT ph" & ChrW(7843) & "i n" & ChrW(7897) & "p"
    ws.Range("H3").Value = "Doanh thu t" & ChrW(237) & "nh thu" & ChrW(7871)
    ws.Range("I3").Value = "T" & ChrW(7927) & " l" & ChrW(7879) & " thu" & ChrW(7871) & " TNDN (%)"
    ws.Range("J3").Value = "Thu" & ChrW(7871) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c mi" & ChrW(7877) & "n, gi" & ChrW(7843) & "m"
    ws.Range("K3").Value = "Thu" & ChrW(7871) & " ph" & ChrW(7843) & "i n" & ChrW(7897) & "p"
    '--- Hàng 4: Ch? s? ---
    Dim arr As Variant
    arr = Array( _
        "41:ct1|838:ct1b", _
        "", _
        "", _
        "41:ct4|838:ct5", _
        "41:ThueGTGT/ct6|838:ThueGTGT/ct7", _
        "41:ThueGTGT/ct7|838:ThueGTGT/ct8", _
        "ThueGTGT/ct9", _
        "ThueTNDN/ct10", _
        "ThueTNDN/ct11", _
        "ThueTNDN/ct12", _
        "ThueTNDN/ct13", _
        "ct14")
    ws.Range("A4:L4").Value = arr
    '--- Dòng 5: BK tham chi?u ---
    ws.Range("A5").Value = "BKThueNTNN"
    ws.Range("D5").Value = "41:/tong_ct6|838:/tong_ct7"
    ws.Range("G5").Value = "tong_ct9"
    ws.Range("H5").Value = "tong_ct10"
    ws.Range("J5").Value = "tong_ct12"
    ws.Range("K5").Value = "tong_ct13"
    ws.Range("L5").Value = "tong_ct14"
    '--- Ð?nh d?ng ---
    With ws.Range("A1:L5")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With
    ws.Columns("A:L").AutoFit
    Rows("4:5").Hidden = True
End Sub

