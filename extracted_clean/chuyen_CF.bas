Attribute VB_Name = "chuyen_CF"
Option Explicit
Public Sub Chuyencf(control As IRibbonControl)
    If Not LicenseGate() Then Exit Sub
'
' chuyen_CF Macro
'
'
    Range("Y8:Y15").Select
    Selection.Copy
    Range("AA8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA16").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-8]C:R[-1]C)"
    Range("Y17:Y25").Select
    Selection.Copy
    Range("AA17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA26").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-9]C:R[-1]C)"
    Range("AA26").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-10]C:R[-1]C)"
    Range("AA27").Select
    ActiveWindow.SmallScroll Down:=10
    Range("Y28:Y34").Select
    Selection.Copy
    Range("AA28").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA35").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-7]C:R[-1]C)"
    Range("AA36").Select
    ActiveWindow.SmallScroll Down:=9
    Range("Y37:Y42").Select
    Selection.Copy
    Range("AA37").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA43").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-6]C:R[-1]C)"
    Range("AA44").Select
    ActiveWindow.SmallScroll Down:=2
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C+R[-9]C+R[-18]C"
    Range("Y45:Y46").Select
    Selection.Copy
    Range("AA45").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AA47").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("AA47").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("AA48").Select
    MsgBox ChrW(272) & ChrW(227) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(117) & ChrW(121) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(120) & ChrW(111) & ChrW(110) & ChrW(103) & ChrW(32) & ChrW(67) & ChrW(70) & ChrW(44) & ChrW(32) & ChrW(109) & ChrW(7901) & ChrW(105) & ChrW(32) & ChrW(97) & ChrW(47) & ChrW(99) & ChrW(32) & ChrW(116) & ChrW(105) & ChrW(7871) & ChrW(112) & ChrW(32) & ChrW(116) & ChrW(7907) & ChrW(99) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(117) & ChrW(121) & ChrW(7875) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(104) & ChrW(117) & ChrW(121) & ChrW(7871) & ChrW(116) & ChrW(32) & ChrW(109) & ChrW(105) & ChrW(110) & ChrW(104) & ChrW(32) & ChrW(273) & ChrW(7847) & ChrW(117) & ChrW(32) & ChrW(107) & ChrW(7923)
End Sub

