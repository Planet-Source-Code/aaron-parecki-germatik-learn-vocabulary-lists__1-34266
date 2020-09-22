VERSION 5.00
Begin VB.Form frmUmlautHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Characters"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   780
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmUmlautHelp.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmUmlautHelp"
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

'does anyone know how to stop the beep
'after the umlaut character is pressed?
'since windows doesn't recognize ALT+o
'as a valid keystroke, it beeps, but this
'program does respond to that. The beeps
'can get very irritating
