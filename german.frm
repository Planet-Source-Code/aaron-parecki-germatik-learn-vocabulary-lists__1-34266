VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGermatik 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Germatik"
   ClientHeight    =   3165
   ClientLeft      =   615
   ClientTop       =   810
   ClientWidth     =   4470
   Icon            =   "german.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3165
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialogList 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".lst"
      DialogTitle     =   "Open List"
      Filter          =   "Germatik Word Lists (*.lst)|*.lst"
   End
   Begin MSComDlg.CommonDialog CommonDialogTest 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".gtk"
      Filter          =   "Germatik Tests (*.gtk)|*.gtk"
   End
   Begin VB.CommandButton btnStartEnd 
      Caption         =   "&Start Test"
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Click to start the test after a list has been loaded"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton btnGuess 
      Caption         =   "Guess"
      Height          =   360
      Left            =   2880
      Picture         =   "german.frx":0E42
      TabIndex        =   2
      ToolTipText     =   "Click to submit your guess (you can also press enter)"
      Top             =   2475
      Width           =   1095
   End
   Begin VB.TextBox guess 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lblRemaining 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Line lineStreak 
      BorderColor     =   &H00FFFFC0&
      X1              =   2640
      X2              =   4200
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label lblStreak 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Streak"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label longestStreakDisp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      ToolTipText     =   "The most number of words gotten right in a row"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblLongest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Longest"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label streakDisp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      ToolTipText     =   "The current number of words gotten right in a row"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label txtPercent 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "The percentage of words asked that were guessed correctly"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Percent"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label txtWrong 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "The number of words guessed incorrectly"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblWrong 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wrong"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblCorrect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Correct"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label txtCorrect 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "The number of words guessed correctly"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label numWordsd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label engDisplay 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenList 
         Caption         =   "&Open List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListEdit 
         Caption         =   "Open List&Editor"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuOpenWordFinder 
         Caption         =   "Open &WordFinder"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuArticle 
         Caption         =   "Open &ArticleFinder"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuStart 
         Caption         =   "Start &Test"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End Test"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "View &High Scores"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Test"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Test &As"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Test"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Configure"
      Begin VB.Menu mnuSelectMode 
         Caption         =   "&Configure Options"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "&Display"
      Begin VB.Menu mnuShowPhoto 
         Caption         =   "Show &Photo in Background"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLargeSize 
         Caption         =   "&Large Size"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuUmlaut 
         Caption         =   "&Special Characters (ü,ß)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Germatik"
      End
   End
End
Attribute VB_Name = "frmGermatik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Germatik version 3.10

'This is a fully functional program designed to help
'in the study of vocabulary lists in German.
'I wrote this program originally two years ago in the
'text-based GW-BASIC. I was doing poorly on my vocabulary
'tests, but after using it, I passed every test with
'flying colors.
'I started this VB version when I was just learning
'Visual Basic, so excuse the sloppiness and/or lack of
'comments in some parts.
'
'I do not claim credit for ALL the code used in this
'program. I wrote all the structural code, and much
'of the code that does the work. But when I encountered
'something I had no idea how to do, I turned to PSC for
'my answers. Thanks to the many people on PSC whose code I
'used. I did modify much of it to suit my needs. If you
'see your code, send me your name and I will include it
'in the source code if you want. Thanks again.
'
'If you have any suggestions, please email me.
'ap@aaronparecki.com
'
'If you modify this program in any way, please send
'me your modifications. I am very interesed to see
'what people do with this.
'
'I would appreciate if you reported any bugs you may
'encounter to me. Feel free to write me!

Public total As Integer 'total asked
Public score As Integer 'total correct
Public ger As String 'string to hold german word
Public eng As String 'string to hold english word
Public recnum As Integer 'record number of chosen word
Public totNum As Integer 'total number of words in list (doean't change during test) Used to set stoopping point
Dim testIsFromSaved As Boolean
Public gerList As New Collection 'holds all the german words
Public engList As New Collection 'holds all the english words
Public gerList1 As New Collection 'holds all the german words
Public engList1 As New Collection 'holds all the english words
Dim getfrom0 As Boolean 'determines which collection to get words from
Dim streak As Integer 'number correct in a row
Dim longestStreak As Integer 'highest number correct in a row
Public askEnglish As Boolean
Public checkGrammar As Boolean 'ask english or german words
Public wrongGer As New Collection
Public wrongEng As New Collection

'variables added when I added different modes
Public TestLimit As Integer
Public ModeIsPractice As Boolean
Public doRandom As Boolean
Public selectedNumWords As Integer 'this is always only the value in the text box on the frmTestMode
Public optAll As Boolean 'if mode is practice, this is true. If mode is test, this is true or false

'score keeping variables
Dim scoPercent As Integer
Dim scoStreak As Integer
Dim scoAllRight As Integer
Dim scoWrong As Integer
Public scoFinalScore As Integer
Public viewScoreOnly As Boolean


Sub doTest()
Dim checkScoreOrTotal As Integer
    
    If ModeIsPractice Then
        checkScoreOrTotal = score
    Else
        checkScoreOrTotal = total
    End If
    
    If checkScoreOrTotal < TestLimit Then
        lblRemaining.Caption = TestLimit - checkScoreOrTotal & " remaining"
        Call getwords
        engDisplay.Caption = eng
        btnGuess.Enabled = True
        guess.Enabled = True
        guess.Text = ""
        If frmGermatik.Enabled = True Then guess.SetFocus
    Else
        'completion of test
        
    'calculate score
        'all scores are percentage of number of words in list
        'percentage
        scoPercent = Int(total * (score / total))
        'longest streak compared to total
        scoStreak = Int(total * (longestStreak / total))
        'all right and none wrong
        If longestStreak = total Then
            scoAllRight = Int(total / 3)
        Else
            scoAllRight = 0
        End If
        'add up score
        scoFinalScore = scoPercent + scoStreak + scoAllRight
        'subtract points for getting word wrong more than once
        scoFinalScore = scoFinalScore - scoWrong
        'multiply by 10 to get final score
        scoFinalScore = scoFinalScore * 10

        engDisplay.Caption = ""
        guess.Text = ""
        
        frmScore.viewOnly = False
        frmScore.Show vbModal
        
        Call menuInitsEnd
        Call endTest
    
    End If
End Sub

Private Sub btnGuess_Click()
    btnGuess.Enabled = False

'****** here is where to add code to ******
'****** analyze guess                ******
    If LTrim(RTrim(guess.Text)) = ger Then
        ' guess is EXACTLY the same
        Call correct
    Else
        If checkGrammar Then
            'If the noun isn't capitalized, pop a message box
            If Len(guess.Text) >= 5 Then 'only check for a noun if guess is longer than 5 chars
                If isNoun(ger) Then
                    If isNoun(guess.Text) Then
                        If Not isLtrCapital(LTrim(guess.Text), 5) Then
                            MsgBox "Don't forget to capitalize your nouns!", vbExclamation + vbOKOnly, "Always Capitalize the Nouns!"
                        End If
                    End If
                End If
            End If
        End If
        
        Call wrong
    End If
'******************************************
'******************************************
    
    total = total + 1
    Call updateStats
    
    Call doTest
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuOpenList.Enabled Then
        If Right(Data.Files(1), 4) = ".lst" Then
            CommonDialogList.fileName = Data.Files(1)
            Call doOpen
            Call mnuStart_Click
        End If
    End If
End Sub

Private Sub guess_KeyDown(KeyCode As Integer, Shift As Integer)
    'If it is possible to guess, on enter, hit guess button
    If btnGuess.Enabled = True Then
        If KeyCode = 13 Then
            Call btnGuess_Click
            Exit Sub
        End If
    End If
    
    'umlaut chars
    Call specialChars(guess, KeyCode, Shift)
End Sub

Private Sub btnGuess_KeyDown(KeyCode As Integer, Shift As Integer)
Dim charToAdd As String
    If Shift > 1 Then Exit Sub
    charToAdd = Chr(CLng(KeyCode))
    
    If KeyCode < 48 Then
        Exit Sub
    End If
    
    If Shift = 0 Then charToAdd = LCase(charToAdd)
    guess = guess & charToAdd

    guess.SetFocus
    guess.SelStart = Len(guess)
End Sub

Sub updateStats()
    txtCorrect.Caption = score
    txtWrong.Caption = total - score
    If total = 0 Then totalR = 1 Else totalR = total
    txtPercent.Caption = Int(score / totalR * 100)
    streakDisp.Caption = streak
    If streak > longestStreak Then
        longestStreak = streak
    End If
    longestStreakDisp.Caption = longestStreak
End Sub

Sub getwords()
    If gerList1.Count = 0 Then getfrom0 = True
    If gerList.Count = 0 Then getfrom0 = False
    
    If doRandom Then
        Do While chooseWordRandom() = False
        Loop
    Else
        If getfrom0 Then
            ger = gerList.Item(1)
            eng = engList.Item(1)
        Else
            ger = gerList1.Item(1)
            eng = engList1.Item(1)
        End If
    End If
End Sub
Function chooseWordRandom() As Boolean
Static lastWord As String 'holds the last word asked (so that the same word doesn't appear twice in a row)
    
    Randomize
    
    If getfrom0 Then
        recnum = Int((gerList.Count) * Rnd) + 1
        ger = gerList.Item(recnum)
        eng = engList.Item(recnum)
    Else
        recnum = Int((gerList1.Count) * Rnd) + 1
        ger = gerList1.Item(recnum)
        eng = engList1.Item(recnum)
    End If
    
    chooseWordRandom = True

End Function

Sub correct()
    score = score + 1
    frmCorrect.Show vbModal
    streak = streak + 1
    
    If getfrom0 Then
        gerList.Remove (recnum)
        engList.Remove (recnum)
    Else
        gerList1.Remove (recnum)
        engList1.Remove (recnum)
    End If
End Sub
Sub wrong()
Dim str As String
    
    streak = 0

    ' if user didn't enter anything, show dialog
    If (guess.Text = "") Then
        Randomize
        If Int(Rnd * 4) = 1 Then
'            Select Case Int(Rnd * 4)
'                Case 0:
                    str = "You're not even going to try?!"
'                Case 1:
'                   could add more messages here
'            End Select
            MsgBox str, vbQuestion + vbOKOnly, "Oops!"
        End If
    End If

    'write guessed word and correct word to collections
    Call collAddItem(wrongGer, ger, -1, True)
    Call collAddItem(wrongEng, eng, -1, True)

    frmWrong.Show vbModal

    If ModeIsPractice Then
        If getfrom0 Then
            gerList1.Add (gerList.Item(recnum))
            engList1.Add (engList.Item(recnum))
            gerList.Remove (recnum)
            engList.Remove (recnum)
        Else 'get from 1, add to 0
            gerList.Add (gerList1.Item(recnum))
            engList.Add (engList1.Item(recnum))
            gerList1.Remove (recnum)
            engList1.Remove (recnum)
        End If
    Else
        gerList.Remove (recnum)
        engList.Remove (recnum)
    End If

End Sub

Sub endTest()
    If Not wrongGer.Count = 0 Then
        frmWrongWords.Show vbModal
        For i = 1 To wrongGer.Count
            wrongGer.Remove (1)
            wrongEng.Remove (1)
        Next i
    End If
    
    total = 0
    score = 0
    streak = 0
    longestStreak = 0
    btnGuess.Enabled = False
    guess.Enabled = False
    engDisplay.Caption = ""
    guess.Text = ""
    lblRemaining.Caption = ""
    If testIsFromSaved Then
        btnStartEnd.Enabled = False
        mnuStart.Enabled = False
        mnuEnd.Enabled = False
        numWordsd.Caption = ""
    End If
    recnum = 1
    testIsFromSaved = False
    If (Not ModeIsPractice) And (Not doall) Then
        TestLimit = selectedNumWords
    Else
        TestLimit = totNum
    End If
End Sub
Sub startTestPrep()
    btnGuess.Enabled = True
    guess.Enabled = True
    
    If Not testIsFromSaved Then
        If Not doesFileExist(CommonDialogList.fileName) Then Exit Sub
        'clear lists in memory
        For i = 1 To gerList.Count
            gerList.Remove (1)
            engList.Remove (1)
        Next
        For i = 1 To gerList1.Count
            gerList1.Remove (1)
            engList1.Remove (1)
        Next
    
        Open CommonDialogList.fileName For Input As #1
        If askEnglish Then
            Do While Not EOF(1)
                Input #1, tempGer
                gerList.Add (tempGer)
                Input #1, tempEng
                engList.Add (tempEng)
            Loop
        Else ' ask german
            Do While Not EOF(1)
                Input #1, tempEng
                engList.Add (tempEng)
                Input #1, tempGer
                gerList.Add (tempGer)
            Loop
        End If
        Close #1
        longestStreakDisp.Caption = 0
    Else
        longestStreakDisp.Caption = longestStreak
        streakDisp.Caption = streak
    End If

    If Not ModeIsPractice Then
        If Not doall Then
            TestLimit = selectedNumWords
        Else
            TestLimit = gerList.Count
        End If
    Else
        TestLimit = gerList.Count
    End If
    
    Call updateStats
End Sub

Sub menuInitsStart()
    mnuEnd.Enabled = True
    mnuStart.Enabled = False
    mnuSaveAs.Enabled = True
    mnuLoad.Enabled = False
    btnStartEnd.Enabled = True
    btnStartEnd.Caption = "End Test"
    mnuListEdit.Enabled = False
    mnuOpenList.Enabled = False
    mnuOpenWordFinder.Enabled = False
    mnuArticle.Enabled = False
    mnuSave.Enabled = testIsFromSaved
    mnuMode.Enabled = False
    mnuHighScores.Enabled = False
End Sub
Sub menuInitsEnd()
    mnuEnd.Enabled = False
    mnuStart.Enabled = True
    mnuSaveAs.Enabled = False
    mnuSave.Enabled = False
    mnuLoad.Enabled = True
    btnStartEnd.Enabled = True
    btnStartEnd.Caption = "Start Test"
    mnuListEdit.Enabled = True
    mnuOpenList.Enabled = True
    mnuOpenWordFinder.Enabled = True
    mnuArticle.Enabled = True
    mnuMode.Enabled = True
    mnuHighScores.Enabled = True
End Sub

Private Sub Exit_Click()
    Unload Me
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Load()
    If getColorDepth < 16 Then
        frmGermatik.Picture = frm256.Picture
    Else
        frmGermatik.Picture = frm24.Picture
    End If
    recnum = 1
    test = 0
    testIsFromSaved = False
    mnuEnd.Enabled = False
    mnuStart.Enabled = False
    btnStartEnd.Enabled = False
    btnGuess.Enabled = False
    guess.Enabled = False
    mnuMode.Enabled = False
    mnuOpenWordFinder.Enabled = False
    mnuArticle.Enabled = False
    
    'Set default test mode: PRACTICE, RANDOM
    ModeIsPractice = True
    doRandom = True
    optAll = True
    askEnglish = True
    checkGrammar = True
    
    Call createFileAssociation
    
End Sub

Private Sub mnuShowPhoto_Click()
    If mnuShowPhoto.Checked = True Then
        ' turn off picture
        mnuShowPhoto.Checked = False
        frmGermatik.Picture = LoadPicture() 'load empty picture
        lblStreak.ForeColor = &H80000012
        lblCurrent.ForeColor = &H80000012
        lblLongest.ForeColor = &H80000012
        lineStreak.BorderColor = &H80000012
        numWordsd.ForeColor = &H80000012
        engDisplay.ForeColor = &H80000012
        lblRemaining.ForeColor = &H80000012
    Else
        ' turn on picture
        mnuShowPhoto.Checked = True
        If getColorDepth < 16 Then
            frmGermatik.Picture = frm256.Picture
        Else
            frmGermatik.Picture = frm24.Picture
        End If
        lblStreak.ForeColor = &HFFFFFF
        lblCurrent.ForeColor = &HFFFFFF
        lblLongest.ForeColor = &HFFFFFF
        lineStreak.BorderColor = &HFFFFC0
        numWordsd.ForeColor = &HFFFFFF
        engDisplay.ForeColor = &HFFFFFF
        lblRemaining.ForeColor = &HFFFFFF
    End If
End Sub

Public Sub mnuStart_Click()
    Call menuInitsStart
    Call startTestPrep
    Call doTest
End Sub

Private Sub mnuEnd_Click()
    Call menuInitsEnd
    Call endTest
End Sub

Private Sub btnStartEnd_Click()
    If mnuStart.Enabled = True Then
        Call menuInitsStart
        Call startTestPrep
        Call doTest
    Else
        Call menuInitsEnd
        Call endTest
    End If
End Sub

Private Function boolToInt(boolVar As Boolean) As Integer
    boolToInt = boolVar
End Function
Private Function intToBool(intVar As Integer) As Boolean
    intToBool = intVar
End Function

Private Sub mnuListEdit_Click()
    frmListEditor.Show vbModal
    Call doOpen
End Sub

Private Sub mnuOpenList_Click()
Dim fileNameTemp

    fileNameTemp = CommonDialogList.fileName
    CommonDialogList.ShowOpen
    If doOpen = False Then
        CommonDialogList.fileName = fileNameTemp
    End If
End Sub

Public Function doOpen() As Boolean
    If CommonDialogList.fileName = "" Then
        doOpen = False
        Exit Function
    End If
        
    If doesFileExist(CommonDialogList.fileName) Then
        totNum = 0
        Open CommonDialogList.fileName For Input As #1
            Do While Not EOF(1)
                Input #1, tempGer
                Input #1, tempEng
                totNum = totNum + 1
            Loop
        Close #1
        Call displayTitleWithFileName(Me, "Germatik - ", CommonDialogList.fileName, False)
        mnuMode.Enabled = True
        mnuOpenWordFinder.Enabled = True
        mnuArticle.Enabled = True
        mnuStart.Enabled = True
        btnStartEnd.Enabled = True
        If Not testIsFromSaved Then
            total = 0
            score = 0
            Call updateStats
        End If
        numWordsd.Caption = totNum & " words"
        filesOK = True
        doOpen = True
    Else
        MsgBox "File not found: '" & fileNameOnly(CommonDialogList.fileName) & "'", vbCritical + vbOKOnly, "Error"
        doOpen = False
    End If

End Function

'###############################################
' code for large size mode

Private Sub mnuLargeSize_Click()
If mnuLargeSize.Checked = False Then
'do make large size
    mnuShowPhoto.Checked = True
    Call mnuShowPhoto_Click
    
    Call doubleSize(txtCorrect)
    Call doubleSize(txtWrong)
    Call doubleSize(txtPercent)
    Call doubleSize(lblCorrect)
    Call doubleSize(lblWrong)
    Call doubleSize(lblPercent)
    Call doubleSize(btnStartEnd)
    Call doubleSize(lblCurrent)
    Call doubleSize(lblLongest)
    Call doubleSize(lblStreak)
    Call doubleSize(lblRemaining)
    Call doubleSize(btnGuess)
    Call doubleSize(guess)
    Call doubleSize(longestStreakDisp)
    Call doubleSize(streakDisp)
    Call doubleSize(numWordsd)
    Call doubleSize(engDisplay)
    
    lineStreak.BorderWidth = lineStreak.BorderWidth * 3 / 2
    lineStreak.X1 = lineStreak.X1 * 3 / 2
    lineStreak.X2 = lineStreak.X2 * 3 / 2
    lineStreak.Y1 = lineStreak.Y1 * 3 / 2
    lineStreak.Y2 = lineStreak.Y2 * 3 / 2
    
    guess.Height = guess.Height - 195
    
    Me.Width = Me.Width * 3 / 2
    Me.Height = (Me.Height * 3 / 2) - 390
    mnuShowPhoto.Enabled = False
    mnuLargeSize.Checked = True
Else
'do make small size
    Call halveSize(txtCorrect)
    Call halveSize(txtWrong)
    Call halveSize(txtPercent)
    Call halveSize(lblCorrect)
    Call halveSize(lblWrong)
    Call halveSize(lblPercent)
    Call halveSize(btnStartEnd)
    Call halveSize(lblCurrent)
    Call halveSize(lblLongest)
    Call halveSize(lblStreak)
    Call halveSize(btnGuess)
    Call halveSize(lblRemaining)
    Call halveSize(guess)
    Call halveSize(longestStreakDisp)
    Call halveSize(streakDisp)
    Call halveSize(numWordsd)
    Call halveSize(engDisplay)
    
    lineStreak.BorderWidth = lineStreak.BorderWidth * 2 / 3
    lineStreak.X1 = lineStreak.X1 * 2 / 3
    lineStreak.X2 = lineStreak.X2 * 2 / 3
    lineStreak.Y1 = lineStreak.Y1 * 2 / 3
    lineStreak.Y2 = lineStreak.Y2 * 2 / 3
    
    guess.Height = guess.Height - 195
    
    Me.Width = Me.Width * 2 / 3
    Me.Height = (Me.Height * 2 / 3) + 255
    mnuShowPhoto.Checked = False
    Call mnuShowPhoto_Click
    mnuShowPhoto.Enabled = True
    mnuLargeSize.Checked = False
End If
    Me.Hide
    centerForm Me
    Me.Show
End Sub

Sub doubleSize(obj As Object)
    changeSize obj, 3, 2
End Sub

Sub halveSize(obj As Object)
    changeSize obj, 2, 3
End Sub
'end code for large size mode
'############################################

Private Sub mnuSelectMode_Click()
    frmTestMode.Show vbModal
End Sub

Private Sub mnuOpenWordFinder_Click()
    frmWordFinder.Show vbModal
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuUmlaut_Click()
    frmUmlautHelp.Show vbModal
End Sub

Private Sub mnuHighScores_Click()
    frmScore.viewOnly = True
    frmScore.Show vbModal
End Sub

Private Sub createFileAssociation()
Dim programName As String
Dim myfiletype As fileType
    
    myfiletype.ProperName = "GermatikList"
    myfiletype.FullName = "Germatik Word List"
    myfiletype.ContentType = "Text"
    myfiletype.Extension = ".LST"
    myfiletype.Commands.Captions.Add "Open"
    programName = """" & App.Path & "\Germatik " & App.Major & _
        "." & App.Minor & App.Revision & ".exe" & """" & " %1"
    myfiletype.Commands.Commands.Add programName
    myfiletype.IconPath = App.Path & "\Germatik " & _
        App.Major & "." & App.Minor & App.Revision & ".exe"
    myfiletype.IconIndex = 0
    'does anyone know how to make the compiled exe
    'contain more than one icon, so that I could
    'choose index 1 here for a small book in a sheet
    'of paper? ap@aaronparecki.com
    
    CreateExtension myfiletype

End Sub

Private Sub mnuArticle_Click()
    frmFindArticle.Show vbModal
End Sub

