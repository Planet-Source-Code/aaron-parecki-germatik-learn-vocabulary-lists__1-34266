VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtract 
   Caption         =   "Extract Words to File"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "frmExtract.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".lst"
      DialogTitle     =   "Save Selected Words to File"
      Filter          =   "Germatik Word Lists (*.lst)|*.lst"
   End
   Begin VB.ListBox lstEnglish 
      Enabled         =   0   'False
      Height          =   2205
      ItemData        =   "frmExtract.frx":000C
      Left            =   2160
      List            =   "frmExtract.frx":0013
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox lstGerman 
      Height          =   2205
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton btnSaveAs 
      Caption         =   "&Save Selected Words"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHelp_Click()
    frmExtractHelp.Show vbModal
End Sub

''''''''''''''''''''''''''''
' ties list boxes together '
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
Private Sub lstGerman_Scroll()
    lstEnglish.TopIndex = lstGerman.TopIndex
End Sub
Private Sub lstEnglish_Scroll()
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub
Private Sub lstGerman_Click() 'for arrow keys
    Call lstGerman_MouseDown(0, 0, 0, 0)
End Sub
Private Sub lstEnglish_Click()
    Call lstEnglish_MouseDown(0, 0, 0, 0)
End Sub
'                          '
''''''''''''''''''''''''''''

Private Sub Form_Load()
    Me.Icon = frmListEditor.Icon
    frmListEditor.Hide
    Call doOpen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmListEditor.Show
End Sub

Private Sub btnSaveAs_Click()
    Call doSave
End Sub

Private Sub doOpen()
    
    lstGerman.clear
    lstEnglish.clear
    
    For i = 0 To frmListEditor.lstGerman.ListCount - 1
        lstGerman.AddItem frmListEditor.lstGerman.List(i)
        lstEnglish.AddItem frmListEditor.lstEnglish.List(i)
    Next

End Sub

Private Sub doSave()
CommonDialog1.ShowSave
If Not CommonDialog1.fileName = "" Then
    Open CommonDialog1.fileName For Output As #1
        For i = 0 To lstGerman.ListCount - 1
            If lstGerman.Selected(i) Then
                Print #1, lstGerman.List(i)
                Print #1, lstEnglish.List(i)
            End If
        Next
    Close #1
    MsgBox "Words successfully saved", vbInformation + vbOKOnly, "Saved"
    Unload Me
End If
End Sub

