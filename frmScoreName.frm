VERSION 5.00
Begin VB.Form frmScoreName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Score!"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "frmScoreName.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "Please enter your name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmScoreName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    If Not LTrim(RTrim(txtName)) = "" Then
        frmScore.newName = txtName
        Unload Me
    Else
        MsgBox "Please enter your name", vbExclamation + vbOKOnly, "No name"
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call btnOK_Click
        Exit Sub
    End If
    If KeyCode = 188 Then
        MsgBox "Not a valid character", vbExclamation + vbOKOnly, "Error"
        txtName = removeStr(txtName, ",")
    End If
End Sub
