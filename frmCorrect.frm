VERSION 5.00
Begin VB.Form frmCorrect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correct!"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1710
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   1710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   315
      Left            =   435
      TabIndex        =   0
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "That's correct!"
      Height          =   255
      Left            =   128
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
End Sub
