VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWrongWords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wrong Words"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmWrongWords.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3045
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".lst"
      DialogTitle     =   "Save As"
      Filter          =   "Germatik Word Lists (*.lst)|*.lst"
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save to File"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3045
      Width           =   1335
   End
   Begin VB.ListBox lstGerman 
      Height          =   2595
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstEnglish 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblGerman 
      Alignment       =   2  'Center
      Caption         =   "German"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblEnglish 
      Alignment       =   2  'Center
      Caption         =   "English"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmWrongWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSave_Click()
    CommonDialog1.ShowSave
    If Not (CommonDialog1.fileName = "") Then
        Open CommonDialog1.fileName For Output As #1
        For i = 0 To lstGerman.ListCount - 1
            ger = lstGerman.List(i)
            eng = lstEnglish.List(i)
            Print #1, ger
            Print #1, eng
        Next i
        Close #1
        MsgBox "List saved successfully", vbInformation + vbOKOnly, "Success"
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
    For i = 1 To frmGermatik.wrongGer.Count
        lstGerman.AddItem frmGermatik.wrongGer.Item(i)
        lstEnglish.AddItem frmGermatik.wrongEng.Item(i)
    Next i
End Sub

Private Sub lstGerman_Scroll()
    lstEnglish.TopIndex = lstGerman.TopIndex
End Sub

Private Sub lstEnglish_Scroll()
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub

Private Sub lstGerman_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lstGerman_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lstGerman_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstEnglish.ListIndex = lstGerman.ListIndex
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub

Private Sub lstEnglish_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lstEnglish_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lstEnglish_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstGerman.ListIndex = lstEnglish.ListIndex
    lstEnglish.TopIndex = lstGerman.TopIndex
End Sub

Private Sub Form_Resize()
    Me.Width = 4770
    If Me.Height >= 2520 Then
        lstGerman.Height = Me.Height - 1260
        lstEnglish.Height = lstGerman.Height
        btnSave.Top = Me.Height - 810
        btnClose.Top = btnSave.Top
    Else
        Me.Height = 2520
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

