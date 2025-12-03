Public SelectedMode As String
Private Sub lblTitle_Click()
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = "GTGT"
    Me.Width = 300
    Me.Height = 180
    Me.BackColor = RGB(245, 245, 245)
    With lblTitle
        .Caption = ChrW(76) & ChrW(7845) & ChrW(121) & ChrW(32) & ChrW(100) & ChrW(7919) & ChrW(32) & ChrW(108) & ChrW(105) & ChrW(7879) & ChrW(117) & ChrW(32) & ChrW(116) & ChrW(7915) & ChrW(58)
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(30, 30, 30)
        .BackStyle = fmBackStyleTransparent
        .Left = 30
        .Top = 20
        .Width = 240
        .Height = 24
    End With
    With cmdFolder
        .Caption = "Folder"
        .Font.Size = 10
        .Font.Bold = False
        .BackColor = RGB(0, 120, 215)
        .ForeColor = vbWhite
        .Width = 100
        .Height = 30
        .Left = 40
        .Top = 70
    End With
    With cmdFile
        .Caption = "File"
        .Font.Size = 10
        .Font.Bold = False
        .BackColor = RGB(0, 120, 215)
        .ForeColor = vbWhite
        .Width = 100
        .Height = 30
        .Left = 150
        .Top = 70
    End With
End Sub
Private Sub cmdFile_Click()
    SelectedMode = "file"
    Me.Hide
End Sub
Private Sub cmdFolder_Click()
    SelectedMode = "folder"
    Me.Hide
End Sub