VERSION 5.00
Begin VB.Form frmArticle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Article"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1995
   Icon            =   "frmArticle.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   1995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.ListBox lstArticle 
      Height          =   645
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Icon = frmListEditor.Icon
    lstArticle.AddItem "der"
    lstArticle.AddItem "die"
    lstArticle.AddItem "das"
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmListEditor.article = lstArticle.List(lstArticle.ListIndex)
End Sub

Private Sub lstArticle_DblClick()
    Unload Me
End Sub
