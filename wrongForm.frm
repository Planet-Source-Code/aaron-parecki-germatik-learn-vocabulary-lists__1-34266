VERSION 5.00
Begin VB.Form frmWrong 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wrong"
   ClientHeight    =   2145
   ClientLeft      =   1560
   ClientTop       =   2340
   ClientWidth     =   4320
   Icon            =   "wrongForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4320
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox inputBox 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1725
      Width           =   1935
   End
   Begin VB.CommandButton OKbutton 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please type in the correct answer below"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label dispGuess 
      Caption         =   "ein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label dispGer 
      Caption         =   "eins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label dispEng 
      Caption         =   "one"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblYourGuess 
      Alignment       =   1  'Right Justify
      Caption         =   "Your Guess:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblGerman 
      Alignment       =   1  'Right Justify
      Caption         =   "Correct German Word"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblEnglish 
      Alignment       =   1  'Right Justify
      Caption         =   "English Word"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmWrong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim localGer As String

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
    
    If frmGermatik.mnuLargeSize.Checked = True Then
        makeLargeSize
    End If
    
    If frmGermatik.askEnglish Then
        lblEnglish.Caption = "English Word"
        lblGerman.Caption = "Correct German Word"
    Else
        lblEnglish.Caption = "German Word"
        lblGerman.Caption = "Correct English Word"
    End If
    dispEng.Caption = frmGermatik.eng
    dispGer.Caption = frmGermatik.ger
    dispGuess.Caption = LTrim(RTrim(frmGermatik.guess))
    inputBox.Enabled = True
End Sub

Private Sub OKbutton_Click()
    If LTrim(RTrim(inputBox.Text)) = frmGermatik.ger Then
'        Me.Hide
'        frmCorrect.Show vbModal
        Unload Me
    Else
        MsgBox "Please type: " & frmGermatik.ger
        inputBox.Text = ""
        inputBox.SetFocus
    End If
End Sub

Private Sub inputBox_KeyDown(KeyCode As Integer, Shift As Integer)
    If OKbutton.Enabled = True Then
        If KeyCode = 13 Then
            Call OKbutton_Click
            Exit Sub
        End If
    End If
    
    Call specialChars(inputBox, KeyCode, Shift)

End Sub

Private Sub makeLargeSize()
    For Each obj In frmWrong
        changeSize obj, 3, 2
    Next
    inputBox.Height = inputBox.Height - 195
    
    Me.Width = Me.Width * 3 / 2
    Me.Height = Me.Height * 3 / 2
End Sub
