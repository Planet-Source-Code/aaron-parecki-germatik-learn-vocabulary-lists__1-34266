VERSION 5.00
Begin VB.Form frmExtractHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help on Extracting Words"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   Icon            =   "frmExtractHelp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblContent 
      Caption         =   "Label1"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmExtractHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmListEditor.Icon
    lblContent.Caption = _
    "With this program you can save certain " & _
    "words from a list to another file. Select " & _
    "consecutive words by clicking the first " & _
    "word you want, then holding down " & _
    "SHIFT and clicking on the last. " & _
    "You can also select single words by " & _
    "holding down CTRL and clicking."
End Sub
