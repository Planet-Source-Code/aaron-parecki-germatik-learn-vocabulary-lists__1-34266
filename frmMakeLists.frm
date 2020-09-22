VERSION 5.00
Begin VB.Form frmListEditor 
   Caption         =   "ListEditor"
   ClientHeight    =   2895
   ClientLeft      =   735
   ClientTop       =   930
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4575
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnChange 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add Words"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtEnglish 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtGerman 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ListBox lstEnglish 
      Height          =   1425
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstGerman 
      Height          =   1425
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblEnglish 
      Alignment       =   2  'Center
      Caption         =   "English"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblGerman 
      Alignment       =   2  'Center
      Caption         =   "German"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New List"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open List"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save List As"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save List"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit ListEditor"
      End
   End
End
Attribute VB_Name = "frmListEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public listEditGerFile As String
Public listEditEngFile As String

Private Sub btnAdd_Click()
    lstGerman.AddItem (txtGerman.Text)
    lstEnglish.AddItem (txtEnglish.Text)
    txtGerman.Text = ""
    txtEnglish.Text = ""
    txtGerman.SetFocus
End Sub

Private Sub btnChange_Click()
Dim gerColl As New Collection
Dim engColl As New Collection
    
    For i = 0 To lstGerman.ListIndex - 1
        gerColl.Add (lstGerman.List(i))
        engColl.Add (lstEnglish.List(i))
    Next i
    gerColl.Add (txtGerman.Text)
    engColl.Add (txtEnglish.Text)
    For i = lstGerman.ListIndex + 1 To lstGerman.ListCount - 1
        gerColl.Add (lstGerman.List(i))
        engColl.Add (lstEnglish.List(i))
    Next i
    
    lstGerman.Clear
    lstEnglish.Clear
    
    For i = 1 To gerColl.Count
        lstGerman.AddItem (gerColl.Item(i))
        lstEnglish.AddItem (engColl.Item(i))
    Next i
    
    txtGerman.Text = ""
    txtEnglish.Text = ""
    
    txtGerman.SetFocus
End Sub

Private Sub btnRemove_Click()
    If lstGerman.ListIndex > -1 Then
        lstGerman.RemoveItem (lstGerman.ListIndex)
        lstEnglish.RemoveItem (lstEnglish.ListIndex)
        txtGerman.Text = ""
        txtEnglish.Text = ""
        txtGerman.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.gerFileList.Refresh
    Form1.engFileList.Refresh
End Sub

Private Sub lstGerman_Click()
    lstEnglish.ListIndex = lstGerman.ListIndex
    txtGerman.Text = lstGerman.List(lstGerman.ListIndex)
    txtEnglish.Text = lstEnglish.List(lstEnglish.ListIndex)
End Sub

Private Sub lstEnglish_Click()
    lstGerman.ListIndex = lstEnglish.ListIndex
    txtGerman.Text = lstGerman.List(lstGerman.ListIndex)
    txtEnglish.Text = lstEnglish.List(lstEnglish.ListIndex)
End Sub

Private Sub txtGerman_Change()
    If txtGerman.Text = "" Or txtEnglish.Text = "" Then
        btnAdd.Enabled = False
        btnChange.Enabled = False
    Else
        btnAdd.Enabled = True
        btnChange.Enabled = True
    End If
End Sub

Private Sub txtEnglish_Change()
    Call txtGerman_Change
End Sub

Private Sub txtGerman_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtEnglish.SetFocus

    pos = txtGerman.SelStart
    
    If Shift = 3 Then
        ' Caps
        If KeyCode = 65 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(196) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' A 196
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 79 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(214) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' O
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 85 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(220) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' U
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 83 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(223) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' SS
            txtGerman.SelStart = pos + 1
        End If
    ElseIf Shift = 2 Then
        'lowercase
        If KeyCode = 65 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(228) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' a
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 79 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(246) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' o
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 85 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(252) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' u
            txtGerman.SelStart = pos + 1
        ElseIf KeyCode = 83 Then
            txtGerman.Text = Left(txtGerman.Text, pos) + Chr(223) + Right(txtGerman.Text, Len(txtGerman.Text) - pos)  ' ss
            txtGerman.SelStart = pos + 1
        End If
    End If

End Sub

Private Sub txtenglish_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call btnAdd_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    listEditGerFile = ""
    listEditEngFile = ""
    lstGerman.Clear
    lstEnglish.Clear
    txtGerman.Text = ""
    txtEnglish.Text = ""
    mnuSave.Enabled = False
End Sub

Private Sub mnuOpen_Click()

    fileDlg4Listeditor.dlgModeIsOpen = True
    fileDlg4Listeditor.Show vbModal

    If Not (listEditGerFile = "") Then
        Open listEditGerFile For Input As #1
        Open listEditEngFile For Input As #2
    
        lstGerman.Clear
        lstEnglish.Clear
        
        Do While Not EOF(1)
            Input #1, ger
            Input #2, eng
            lstGerman.AddItem (ger)
            lstEnglish.AddItem (eng)
        Loop
        
        Close #1: Close #2
    End If

End Sub

Function initFiles(localGerFile As String, localEngFile As String) As Boolean
    gerNum = 0
    engnum = 0
    
    initFiles = True
    
    Open localGerFile For Input As #1
    Open localEngFile For Input As #2
    
    Do While Not EOF(1)
        Input #1, tempStr
        gerNum = gerNum + 1
    Loop
    
    Do While Not EOF(2)
        Input #2, tempStr
        engnum = engnum + 1
    Loop
    
    Close #1
    Close #2

    If gerNum <> engnum Then
        MsgBox "Files have different number of records.", vbCritical, "Error!"
        initFiles = False
        Exit Function
    End If

    initFiles = True
End Function

Private Sub mnuSave_Click()
    If Not (listEditGerFile = "" Or listEditEngFile = "") Then
        listEditGerFile = Replace(listEditGerFile, "\\", "\")
        listEditEngFile = Replace(listEditEngFile, "\\", "\")
        Open listEditGerFile For Output As #1
        Open listEditEngFile For Output As #2
        
        For i = 0 To lstGerman.ListCount - 1
            ger = lstGerman.List(i)
            eng = lstEnglish.List(i)
            Print #1, ger
            Print #2, eng
        Next i
        
        Close #1
        Close #2
        
        MsgBox "Files saved successfully", vbOKOnly, "Saved"
    Else
        Call mnuSaveAs_click
    End If
End Sub

Private Sub mnuSaveAs_click()
    
    fileDlg4Listeditor.dlgModeIsOpen = False
    fileDlg4Listeditor.Show vbModal
    
    If Not (listEditGerFile = "" Or listEditEngFile = "") Then
        Call mnuSave_Click
        mnuSave.Enabled = True
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Width = 4695
    txtGerman.Top = Me.Height - 1665
    txtEnglish.Top = txtGerman.Top
    lstGerman.Height = Me.Height - 2160
    lstEnglish.Height = lstGerman.Height
    btnRemove.Top = Me.Height - 1185
    btnAdd.Top = btnRemove.Top
    btnChange.Top = btnRemove.Top
End Sub

