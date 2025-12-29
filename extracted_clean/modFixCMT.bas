Attribute VB_Name = "modFixCMT"
Option Explicit

Public Sub FixCMT()
    If Not ConfirmActiveSheetRisk("Vi tri comment se duoc can chinh lai.") Then Exit Sub
    Dim o As Comment
    For Each o In ActiveSheet.Comments
        o.Shape.Top = o.Parent.Top
    Next
End Sub

Public Sub FixCMT_Ribbon(control As IRibbonControl)
    FixCMT
End Sub
