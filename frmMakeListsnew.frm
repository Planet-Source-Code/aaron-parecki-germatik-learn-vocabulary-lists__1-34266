VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListEditor 
   Caption         =   "ListEditor"
   ClientHeight    =   2895
   ClientLeft      =   735
   ClientTop       =   930
   ClientWidth     =   4575
   Icon            =   "frmMakeListsnew.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".lst"
      Filter          =   "Germatik Word Lists (*.lst)|*.lst"
   End
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
      ItemData        =   "frmMakeListsnew.frx":0E42
      Left            =   240
      List            =   "frmMakeListsnew.frx":0E44
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblNumWords 
      Alignment       =   2  'Center
      Caption         =   "0 words"
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEnglish 
      Alignment       =   1  'Right Justify
      Caption         =   "English"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblGerman 
      Caption         =   "German"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New List"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add &List to Current"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save List &As"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save List"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit ListEditor"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmListEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public article As String
Dim fileNameSaved As String 'filename of opened file, used
                            'to reopen file if blank list
                            'is not saved

Private Sub btnAdd_Click()
Dim pos As Integer

    Call askToCapitalizeNoun
    
    ' add the word to the list only if it doesn't exist,
    ' removing any spaces from beg or end
    If (Not wordIsInList(lstGerman, LTrim(RTrim(txtGerman.Text)), True)) And (Not wordIsInList(lstEnglish, LTrim(RTrim(txtEnglish.Text)), True)) Then
        'check to see if the first letter is capitalized.
        ' if it is, then ask if it's a noun.
        ' if so, then ask which article it gets.
        If isLtrCapital(txtGerman.Text, 1) Then
            msg = "Is '" & LTrim(RTrim(txtGerman.Text)) & "' a noun?"
            inputIsNoun = MsgBox(msg, vbQuestion + vbYesNo, "Noun?")
            If inputIsNoun = vbYes Then
                frmArticle.Show vbModal
                txtGerman.Text = article & " " & txtGerman.Text
            End If
        End If
        If lstGerman.ListIndex = -1 Then
            lstGerman.AddItem LTrim(RTrim(txtGerman.Text))
            lstEnglish.AddItem LTrim(RTrim(txtEnglish.Text))
            lstGerman.ListIndex = lstGerman.ListCount - 1
        Else
            pos = lstGerman.ListIndex
            lstGerman.AddItem LTrim(RTrim(txtGerman.Text)), pos + 1
            lstEnglish.AddItem LTrim(RTrim(txtEnglish.Text)), pos + 1
            lstGerman.ListIndex = pos + 1
        End If
        mnuSave.Enabled = True
        lblNumWords = lstGerman.ListCount & " words"
    Else
        ' both words are in list
        If wordIsInList(lstGerman, LTrim(RTrim(txtGerman.Text)), True) And wordIsInList(lstEnglish, LTrim(RTrim(txtEnglish.Text)), True) Then
            MsgBox "Words already exist in list", vbExclamation + vbOKOnly, "Duplicate Words"
        Else
            ' one word is in list
            If wordIsInList(lstGerman, LTrim(RTrim(txtGerman.Text)), True) Then
                msgLang = "German"
            Else
                msgLang = "English"
            End If
            MsgBox msgLang & " word already in list", vbOKOnly + vbExclamation, "Duplicate Word"
        End If
    End If
    
    txtGerman.Text = ""
    txtEnglish.Text = ""
    txtGerman.SetFocus
    
End Sub

Private Sub btnChange_Click()
'Dim gerColl As New Collection
'Dim engColl As New Collection
Dim pos As Integer
Dim topPos As Integer
    
    'save the position of the lists
    pos = lstGerman.ListIndex
    topPos = lstGerman.TopIndex

    If Form1.checkGrammar Then
        Call askToCapitalizeNoun
    End If
    
    'do the change
    'delete old word, add new word
    changeList lstGerman, LTrim(RTrim(txtGerman))
    changeList lstEnglish, LTrim(RTrim(txtEnglish))
    
    'load the position of the lists and
    'clear the text fields
    lstGerman.ListIndex = pos
    lstEnglish.ListIndex = pos
    lstGerman.TopIndex = topPos
    lstEnglish.TopIndex = topPos
    mnuSave.Enabled = True
    txtGerman.Text = ""
    txtEnglish.Text = ""
    txtGerman.SetFocus
End Sub

Private Sub btnRemove_Click()
Dim pos As Integer
    If lstGerman.ListIndex > -1 Then
        pos = lstGerman.ListIndex
        lstGerman.RemoveItem (pos)
        lstEnglish.RemoveItem (pos)
        txtGerman.SetFocus
        lblNumWords = lstGerman.ListCount & " words"
        mnuSave.Enabled = True
        Do Until pos < lstGerman.ListCount
            pos = pos - 1
        Loop
        lstGerman.ListIndex = pos
    End If
End Sub

Private Sub askToCapitalizeNoun()
Dim newGer As String
    
    If isNoun(LTrim(txtGerman.Text)) Then
        If Not isLtrCapital(LTrim(txtGerman.Text), 5) Then
            newGer = LTrim(txtGerman.Text)
            Mid(newGer, 5, 1) = UCase(Mid(newGer, 5, 1))
            If MsgBox("Do you mean, " & newGer & "?", vbQuestion + vbYesNo, "Capitalize Noun?") = vbYes Then
                txtGerman.Text = newGer
            End If
        End If
    End If

End Sub

Private Sub lstGerman_Click()
    lstEnglish.ListIndex = lstGerman.ListIndex
    txtGerman.Text = lstGerman.List(lstGerman.ListIndex)
    txtEnglish.Text = lstEnglish.List(lstEnglish.ListIndex)
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub

Private Sub lstEnglish_Click()
    lstGerman.ListIndex = lstEnglish.ListIndex
    txtGerman.Text = lstGerman.List(lstGerman.ListIndex)
    txtEnglish.Text = lstEnglish.List(lstEnglish.ListIndex)
    lstEnglish.TopIndex = lstGerman.TopIndex
End Sub

Private Sub lstGerman_Scroll()
    lstEnglish.TopIndex = lstGerman.TopIndex
End Sub

Private Sub lstEnglish_Scroll()
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub

Private Sub txtGerman_Change()
    Call doMenuEnableOrDisable
End Sub

Private Sub txtEnglish_Change()
    Call doMenuEnableOrDisable
End Sub

Private Sub txtGerman_KeyDown(keyCode As Integer, shift As Integer)
    If keyCode = 13 Then txtEnglish.SetFocus
    
    If keyCode = 188 Then
        MsgBox "Not a valid character", vbExclamation + vbOKOnly, "Error"
        txtGerman.Text = removeStr(txtGerman.Text, ",")
    End If
    
    Call specialChars(txtGerman, keyCode, shift)
End Sub

Private Sub txtenglish_KeyDown(keyCode As Integer, shift As Integer)
    If keyCode = 13 Then Call btnAdd_Click

    If keyCode = 188 Then
        MsgBox "Not a valid character", vbExclamation + vbOKOnly, "Error"
        txtEnglish.Text = removeStr(txtEnglish.Text, ",")
    End If

End Sub

Private Sub doMenuEnableOrDisable()
    If txtGerman.Text = "" Or txtEnglish.Text = "" Then
        btnAdd.Enabled = False
        btnChange.Enabled = False
    Else
        btnAdd.Enabled = True
        btnChange.Enabled = True
    End If
End Sub

Private Sub txtGerman_GotFocus()
    txtGerman.SelStart = 0
    txtGerman.SelLength = Len(txtGerman.Text)
End Sub

Private Sub txtEnglish_GotFocus()
    txtEnglish.SelStart = 0
    txtEnglish.SelLength = Len(txtEnglish.Text)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    fileNameSaved = CommonDialog1.fileName
    CommonDialog1.fileName = ""
    Me.Caption = "ListEditor"
    lstGerman.Clear
    lstEnglish.Clear
    txtGerman.Text = ""
    txtEnglish.Text = ""
    lblNumWords = lstGerman.ListCount & " words"
    mnuSave.Enabled = False
End Sub

Private Sub mnuOpen_Click()
Dim fileNameTemp As String
    
    fileNameTemp = CommonDialog1.fileName
    CommonDialog1.fileName = ""
    
    CommonDialog1.ShowOpen
    If Not (CommonDialog1.fileName = "") Then
        lstGerman.Clear
        lstEnglish.Clear
        Call doOpen
    Else
        CommonDialog1.fileName = fileNameTemp
    End If
End Sub

Private Sub mnuAdd_Click()
Dim fileNameTemp As String
    
    fileNameTemp = CommonDialog1.fileName
    CommonDialog1.fileName = ""
    
    CommonDialog1.ShowOpen
    If Not (CommonDialog1.fileName = "") Then
        Call doOpen
        mnuSave.Enabled = True
    End If
    
    CommonDialog1.fileName = fileNameTemp
End Sub

Private Sub doOpen()
    If Not CommonDialog1.fileName = "" Then
        Open CommonDialog1.fileName For Input As #1
            Do While Not EOF(1)
                Input #1, ger
                Input #1, eng
                lstGerman.AddItem (ger)
                lstEnglish.AddItem (eng)
            Loop
        Close #1
        
        Call displayTitleWithFileName(Me, "ListEditor - ", CommonDialogList.fileName)
        lblNumWords = lstGerman.ListCount & " words"
        txtGerman.Text = ""
        txtEnglish.Text = ""
        mnuAdd.Enabled = True
        fileNameSaved = CommonDialog1.fileName
    End If
End Sub

Private Sub mnuSaveAs_click()
    
    CommonDialog1.ShowSave
    
    If Not (CommonDialog1.fileName = "") Then
        Call mnuSave_Click
        mnuSave.Enabled = True
    End If
    
End Sub

Private Sub mnuSave_Click()
    If Not (CommonDialog1.fileName = "") Then
        Open CommonDialog1.fileName For Output As #1
        
        For i = 0 To lstGerman.ListCount - 1
            ger = lstGerman.List(i)
            eng = lstEnglish.List(i)
            Print #1, ger
            Print #1, eng
        Next i
        
        Close #1
        
    Else
        Call mnuSaveAs_click
    End If
End Sub

'############################
'  general form stuff

Private Sub Form_Resize()
    Me.Width = 4695
    If Me.Height >= 2520 Then
        txtGerman.Top = Me.Height - 1665
        txtEnglish.Top = txtGerman.Top
        lstGerman.Height = Me.Height - 2160
        lstEnglish.Height = lstGerman.Height
        btnRemove.Top = Me.Height - 1185
        btnAdd.Top = btnRemove.Top
        btnChange.Top = btnRemove.Top
    Else
        Me.Height = 2520
    End If
End Sub

Private Sub Form_Load()
    mnuSave.Enabled = False

    If Form1.btnStartEnd.Enabled = True Then
        CommonDialog1.fileName = Form1.CommonDialogList.fileName
        Call doOpen
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.CommonDialogList.fileName = CommonDialog1.fileName
    leave = MsgBox("Save list before exit?", vbYesNo + vbExclamation, "Save?")
    Select Case leave
        Case 2: 'cancel
            
        Case 6: 'yes
            Call mnuSave_Click
            If CommonDialog1.fileName = "" Then
                Form1.CommonDialogList.fileName = fileNameSaved
            End If
        Case 7: 'no
            If CommonDialog1.fileName = "" Then
                Form1.CommonDialogList.fileName = fileNameSaved
            End If
    End Select
    Unload Me
End Sub
