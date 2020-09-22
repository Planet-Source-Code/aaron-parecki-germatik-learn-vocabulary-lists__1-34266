VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListEditor 
   Caption         =   "ListEditor"
   ClientHeight    =   2865
   ClientLeft      =   735
   ClientTop       =   930
   ClientWidth     =   4575
   Icon            =   "ListEditor.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
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
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Remove selected words"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add &Words"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Add words in input boxes to the list"
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
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstGerman 
      Height          =   1425
      ItemData        =   "ListEditor.frx":0E42
      Left            =   240
      List            =   "ListEditor.frx":0E44
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblNumWords 
      Alignment       =   2  'Center
      Caption         =   "0 words"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Total number of words in loaded list"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEnglish 
      Alignment       =   1  'Right Justify
      Caption         =   "English"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblGerman 
      Caption         =   "German"
      Height          =   255
      Left            =   360
      TabIndex        =   6
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
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract Words to File"
         Enabled         =   0   'False
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
Dim fileNameTemp As String
Dim fileNameSaved As String 'filename of opened file, used
                            'to reopen file if blank list
                            'is not saved
Dim hasJustBeenSaved As Boolean

Private Sub btnAdd_Click()
Dim pos As Integer

    Call askToCapitalizeNoun
    
    If wordIsOkToAdd Then
        
        If lstGerman.ListIndex = -1 Then
            lstGerman.AddItem Trim(txtGerman.Text)
            lstEnglish.AddItem Trim(txtEnglish.Text)
            lstGerman.ListIndex = lstGerman.ListCount - 1
        Else
            pos = lstGerman.ListIndex
            lstGerman.AddItem Trim(txtGerman.Text), pos + 1
            lstEnglish.AddItem Trim(txtEnglish.Text), pos + 1
            lstGerman.ListIndex = pos + 1
        End If
        mnuSave.Enabled = True
        lblNumWords = lstGerman.ListCount & " words"
        
        mnuSave.Enabled = True
    
    End If
    
    txtGerman.Text = ""
    txtEnglish.Text = ""
    txtGerman.SetFocus

End Sub

'Private Sub btnChange_Click()
''Dim pos As Integer
'Dim topPos As Integer
'
'    'save the position of the lists
'    pos = lstGerman.ListIndex
'    topPos = lstGerman.TopIndex
'
'    Call askToCapitalizeNoun
'
'    'do the change
'    changeList lstGerman, Trim(txtGerman), -1
'    changeList lstEnglish, Trim(txtEnglish), -1
'
'    'load the position of the lists and
'    'clear the text fields
'    lstGerman.ListIndex = pos
'    lstEnglish.ListIndex = pos
'    lstGerman.TopIndex = topPos
'    lstEnglish.TopIndex = topPos
'    mnuSave.Enabled = True
'    txtGerman.Text = ""
'    txtEnglish.Text = ""
'    txtGerman.SetFocus
'End Sub

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
        mnuSave.Enabled = True
    End If
End Sub

Private Function wordIsOkToAdd() As Boolean
Dim response As VbMsgBoxResult
Dim msgWordFrom As String
Dim msgWordDup As String
Dim msgWordTo As String
Dim chgPos As Integer
Dim message As String
Dim germanIsInList As Boolean
    
    'make sure both words are not in list
    If (Not wordIsInList(lstGerman, Trim(txtGerman), True)) And (Not wordIsInList(lstEnglish, Trim(txtEnglish), True)) Then
        'check to see if the first letter is capitalized.
        ' if it is, then ask if it's a noun.
        ' if so, then ask which article it gets.
        If isLtrCapital(txtGerman, 1) Then
            msg = "Is '" & Trim(txtGerman) & "' a noun?"
            inputIsNoun = MsgBox(msg, vbQuestion + vbYesNo, "Noun?")
            If inputIsNoun = vbYes Then
                frmArticle.Show vbModal
                txtGerman = article & " " & txtGerman
            End If
        End If
        ' check to see if the word is a noun, if so, check
        ' if it is already in the list with a different article
        If isNoun(Trim(txtGerman)) Then
            If Not gerWordIsInList(lstGerman, Trim(txtGerman)) Then
                
                'everything passed, ok to add
                wordIsOkToAdd = True
            
            Else
                MsgBox "German word is already in the list with a different article", vbExclamation + vbOKOnly, "Noun already in list"
                wordIsOkToAdd = False
            End If
        Else
            wordIsOkToAdd = True
        End If
    Else
        ' both words are in list
        If wordIsInList(lstGerman, Trim(txtGerman.Text), True) _
          And wordIsInList(lstEnglish, Trim(txtEnglish.Text), True) Then
            MsgBox "Both words already exist in list", vbExclamation + vbOKOnly, "Duplicate Words"
            wordIsOkToAdd = False
        Else
            ' one word is in list
            If wordIsInList(lstGerman, Trim(txtGerman), True) Then
                chgPos = listFindPos(lstGerman, Trim(txtGerman), True)
                msgWordFrom = lstEnglish.List(chgPos)
                msgWordDup = txtGerman
                msgWordTo = txtEnglish
                germanIsInList = True
            Else
                chgPos = listFindPos(lstEnglish, Trim(txtEnglish), True)
                msgWordFrom = lstGerman.List(chgPos)
                msgWordDup = txtEnglish
                msgWordTo = txtGerman
                germanIsInList = False
            End If
            lstGerman.ListIndex = chgPos 'this line changes the contents of the text boxes
            message = "'" & msgWordDup & "' already exists in list." & vbCrLf
            message = message & "Would you like to change '" & msgWordFrom & "' to '" & Trim(msgWordTo) & "?"
            response = MsgBox(message, vbQuestion + vbYesNo, "Change?")
            
            If response = vbYes Then
                'if the german word is in the list...
                If germanIsInList Then
                    changeList lstEnglish, msgWordTo, chgPos
                    lstEnglish.ListIndex = chgPos
                Else 'the english word is in the list
                    changeList lstGerman, msgWordTo, chgPos
                    lstGerman.ListIndex = chgPos
                End If
            End If
            wordIsOkToAdd = False
        End If
    End If

End Function

Private Sub askToCapitalizeNoun()
Dim newGer As String
    
    If isNoun(LTrim(txtGerman)) Then
        If Not isLtrCapital(LTrim(txtGerman), 5) Then
            newGer = LTrim(txtGerman)
            Mid(newGer, 5, 1) = UCase(Mid(newGer, 5, 1))
            If MsgBox("Do you mean, " & newGer & "?", vbQuestion + vbYesNo, "Capitalize Noun?") = vbYes Then
                txtGerman.Text = newGer
            End If
        End If
    End If

End Sub

''''''''''''''''''''''''''''
' ties list boxes together
Private Sub lstGerman_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lstGerman_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lstGerman_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstEnglish.ListIndex = lstGerman.ListIndex
    txtGerman.Text = lstGerman.List(lstGerman.ListIndex)
    txtEnglish.Text = lstEnglish.List(lstEnglish.ListIndex)
    lstGerman.TopIndex = lstEnglish.TopIndex
End Sub
Private Sub lstEnglish_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lstEnglish_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lstEnglish_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub lstGerman_Click() 'for arrow keys
    Call lstGerman_MouseDown(0, 0, 0, 0)
End Sub
Private Sub lstEnglish_Click()
    Call lstEnglish_MouseDown(0, 0, 0, 0)
End Sub
'
''''''''''''''''''''''''''''

Private Sub mnuExtract_Click()
    frmExtract.Show vbModal
End Sub

Private Sub txtGerman_Change()
    Call doMenuEnableOrDisable
End Sub

Private Sub txtEnglish_Change()
    Call doMenuEnableOrDisable
End Sub

Private Sub txtGerman_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtEnglish.SetFocus
    
    If KeyCode = 188 Then
        MsgBox "Not a valid character", vbExclamation + vbOKOnly, "Error"
        txtGerman.Text = removeStr(txtGerman.Text, ",")
    End If
    
    Call specialChars(txtGerman, KeyCode, Shift)
End Sub

Private Sub txtenglish_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call btnAdd_Click

    If KeyCode = 188 Then
        MsgBox "Not a valid character", vbExclamation + vbOKOnly, "Error"
        txtEnglish.Text = removeStr(txtEnglish.Text, ",")
    End If

End Sub

Private Sub doMenuEnableOrDisable()
    If txtGerman.Text = "" Or txtEnglish.Text = "" Then
        btnAdd.Enabled = False
'        btnChange.Enabled = False
    Else
        btnAdd.Enabled = True
'        btnChange.Enabled = True
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
    lstGerman.clear
    lstEnglish.clear
    txtGerman.Text = ""
    txtEnglish.Text = ""
    lblNumWords = lstGerman.ListCount & " words"
    mnuSave.Enabled = False
End Sub

Private Sub mnuOpen_Click()
    
    fileNameTemp = CommonDialog1.fileName
    CommonDialog1.fileName = ""
    
    CommonDialog1.ShowOpen
    If Not (CommonDialog1.fileName = "") Then
        If doOpen(CommonDialog1.fileName, True) Then
            Call displayTitleWithFileName(Me, "ListEditor - ", CommonDialog1.fileName, False)
        Else
            CommonDialog1.fileName = fileNameTemp
        End If
    Else
        CommonDialog1.fileName = fileNameTemp
    End If

End Sub

Private Sub mnuAdd_Click()
    
    fileNameTemp = CommonDialog1.fileName 'save first filename
    CommonDialog1.fileName = "" 'clear filename
    
    CommonDialog1.ShowOpen 'get new filename to open
    If Not (CommonDialog1.fileName = "") Then 'don't try to open if filename is blank
        If doOpen(CommonDialog1.fileName, False) Then
            mnuSave.Enabled = True
            Call displayTitleWithFileName(Me, "ListEditor - ", fileNameTemp, False)
        End If
    End If
    
    CommonDialog1.fileName = fileNameTemp 'set filename back to original

End Sub

Private Function doOpen(filenametoopen As String, clear As Boolean) As Boolean
    If filenametoopen = "" Then
        doOpen = False
        Exit Function
    End If
        
    If doesFileExist(filenametoopen) Then
        If clear Then
            lstGerman.clear
            lstEnglish.clear
        End If
        Open filenametoopen For Input As #1
            Do While Not EOF(1)
                Input #1, ger
                Input #1, eng
                lstGerman.AddItem (ger)
                lstEnglish.AddItem (eng)
            Loop
        Close #1
    
        lblNumWords = lstGerman.ListCount & " words"
        txtGerman.Text = ""
        txtEnglish.Text = ""
        mnuAdd.Enabled = True
        mnuExtract.Enabled = True
        fileNameSaved = filenametoopen
        doOpen = True
    Else
        MsgBox "File not found: '" & fileNameOnly(filenametoopen) & "'", vbCritical + vbOKOnly, "Error"
        doOpen = False
    End If

End Function

Private Sub mnuSaveAs_click()
    
    If Not lstGerman.ListCount = 0 Then
        CommonDialog1.ShowSave
        If Not (CommonDialog1.fileName = "") Then
            Call mnuSave_Click
            mnuSave.Enabled = True
            Call displayTitleWithFileName(Me, "ListEditor - ", CommonDialog1.fileName, False)
        End If
    Else
        MsgBox "Cannot save empty list", vbExclamation + vbOKOnly, "Cannot save"
    End If
    
End Sub

Private Sub mnuSave_Click()
    If Not lstGerman.ListCount = 0 Then
        If Not (CommonDialog1.fileName = "") Then
            Open CommonDialog1.fileName For Output As #1
            For i = 0 To lstGerman.ListCount - 1
                ger = lstGerman.List(i)
                eng = lstEnglish.List(i)
                Print #1, ger
                Print #1, eng
            Next i
            Close #1
            mnuSave.Enabled = False
        Else
            Call mnuSaveAs_click
        End If
    Else
        MsgBox "Cannot save empty list", vbExclamation + vbOKOnly, "Cannot save"
    End If
End Sub

'############################
'  general form stuff

Private Sub Form_Resize()
    Me.WindowState = 0
    Me.Width = 4695
    If Me.Height >= 4000 Then
        txtGerman.Top = Me.Height - 1665
        txtEnglish.Top = txtGerman.Top
        lstGerman.Height = Me.Height - 2160
        lstEnglish.Height = lstGerman.Height
        btnRemove.Top = Me.Height - 1160
        btnAdd.Top = Me.Height - 1160
    Else
        Me.Height = 4000
    End If
End Sub

Private Sub Form_Load()
    If frmGermatik.btnStartEnd.Enabled = True Then
        CommonDialog1.fileName = frmGermatik.CommonDialogList.fileName
        Call doOpen(CommonDialog1.fileName, True)
        Call displayTitleWithFileName(Me, "ListEditor - ", CommonDialog1.fileName, False)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmGermatik.CommonDialogList.fileName = CommonDialog1.fileName
    
    If Not lstGerman.ListCount = 0 Then
        If mnuSave.Enabled = True Then
            leave = MsgBox("Save list before exit?", vbExclamation + vbYesNoCancel, "Save?")
            Select Case leave
                Case 2: 'cancel
                    Cancel = 1 'if cancel is pressed on the message box, then cancel the unload
                Case 6: 'yes
                    Call mnuSave_Click
            End Select
        End If
        If CommonDialog1.fileName = "" Then
            frmGermatik.CommonDialogList.fileName = fileNameSaved
        End If
    Else
        'new list was created, then listEditor was closed
        If CommonDialog1.fileName = "" Then
            frmGermatik.CommonDialogList.fileName = fileNameSaved
        End If
    End If
    
    Unload Me
End Sub
