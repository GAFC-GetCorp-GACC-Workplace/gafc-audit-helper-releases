VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateNKC 
   Caption         =   "UserForm2"
   ClientHeight    =   5100
   ClientLeft      =   375
   ClientTop       =   1500
   ClientWidth     =   34605
   OleObjectBlob   =   "frmCreateNKC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateNKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public SelectedMode As String

Private Sub lblTitle_Click()
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Chon mau so"
    Me.Width = 340
    Me.Height = 150
    Me.BackColor = RGB(245, 245, 245)

    With lblTitle
        .Caption = ChrW(66) & ChrW(7841) & ChrW(110) & ChrW(32) & ChrW(109) & ChrW(117) & ChrW(7889) & ChrW(110) & ChrW(32) & ChrW(116) & ChrW(7841) & ChrW(111) & ChrW(32) & ChrW(108) & ChrW(111) & ChrW(7841) & ChrW(105) & ChrW(32) & ChrW(115) & ChrW(7893) & ChrW(32) & ChrW(110) & ChrW(224) & ChrW(111) & ChrW(63)
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(30, 30, 30)
        .BackStyle = fmBackStyleTransparent
        .Left = 30
        .Top = 20
        .Width = 280
        .Height = 24
    End With

    With cmdFolder
        .Caption = ChrW(84) & ChrW(7841) & ChrW(111) & ChrW(32) & ChrW(115) & ChrW(7893) & ChrW(32) & ChrW(99) & ChrW(104) & ChrW(432) & ChrW(97) & ChrW(32) & ChrW(120) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(253)
        .Font.Size = 10
        .Font.Bold = False
        .BackColor = RGB(0, 120, 215)
        .ForeColor = vbWhite
        .Width = 130
        .Height = 50
        .Left = 30
        .Top = 70
    End With

    With cmdFile
        .Caption = ChrW(84) & ChrW(7841) & ChrW(111) & ChrW(32) & ChrW(115) & ChrW(7893) & ChrW(32) & ChrW(273) & ChrW(227) & ChrW(32) & ChrW(120) & ChrW(7917) & ChrW(32) & ChrW(108) & ChrW(253)
        .Font.Size = 10
        .Font.Bold = False
        .BackColor = RGB(0, 176, 80)
        .ForeColor = vbWhite
        .Width = 130
        .Height = 50
        .Left = 180
        .Top = 70
    End With
End Sub

Private Sub cmdFile_Click()
    SelectedMode = "template"
    Me.Hide
End Sub

Private Sub cmdFolder_Click()
    SelectedMode = "raw"
    Me.Hide
End Sub
