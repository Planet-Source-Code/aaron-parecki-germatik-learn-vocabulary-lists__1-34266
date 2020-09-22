VERSION 5.00
Begin VB.Form frmTestMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure Test"
   ClientHeight    =   3525
   ClientLeft      =   2.45790e5
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmTestMode.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameGrammar 
      Caption         =   "Spelling/Grammar"
      Height          =   1095
      Left            =   1680
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
      Begin VB.CheckBox chkGrammar 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblGrammar 
         Caption         =   "Check for capitalized nouns, etc."
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frameAsk 
      Caption         =   "Ask which list?"
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
      Begin VB.OptionButton optAskGerman 
         Caption         =   "Ask German"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   650
         Width           =   1215
      End
      Begin VB.OptionButton optAskEnglish 
         Caption         =   "Ask English"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Close without saving changes"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Save changes and close"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame frameInfo 
      Caption         =   "Selected Mode"
      Height          =   1575
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.Label lblInfo 
         Caption         =   "Info about selected mode goes here"
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CheckBox chkRandom 
      Caption         =   "Random"
      Height          =   255
      Left            =   140
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton optModeTest 
      Caption         =   "Test Mode"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optModePractice 
      Caption         =   "Practice Mode"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame frameTestModeOptions 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1455
      Begin VB.TextBox txtNumWords 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Text            =   "1"
         Top             =   120
         Width           =   375
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optWords 
         Caption         =   "        Words"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   3480
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   3480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000016&
      BorderStyle     =   0  'Transparent
      X1              =   480
      X2              =   3840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblNumWords 
      Caption         =   "0 words"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTestMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
    If frmGermatik.doRandom Then chkRandom.Value = vbChecked Else chkRandom.Value = vbUnchecked
    optModePractice = frmGermatik.ModeIsPractice
    txtNumWords.Text = frmGermatik.selectedNumWords
    optAll = frmGermatik.optAll
    optWords = Not frmGermatik.optAll
    txtNumWords.Enabled = Not optAll
    If frmGermatik.ModeIsPractice Then Call frameTestModeOptionsDisable
    lblNumWords.Caption = frmGermatik.totNum & " words in list"
    If frmGermatik.askEnglish Then
        optAskEnglish = True
        If frmGermatik.checkGrammar Then chkGrammar.Value = vbChecked Else chkGrammar.Value = vbUnchecked
    Else
        optAskGerman = True
        chkGrammar.Value = vbUnchecked
        chkGrammar.Enabled = False
    End If
End Sub

Private Sub optModePractice_Click()
    Call frameTestModeOptionsDisable
    Call displayInfo
End Sub

Private Sub optModeTest_Click()
    Call frameTestModeOptionsEnable
    Call displayInfo
End Sub

Private Sub chkRandom_Click()
    Call displayInfo
End Sub

Private Sub optWords_Click()
    If txtNumWords.Text > frmGermatik.totNum Then
        txtNumWords.Text = frmGermatik.totNum
    End If
    txtNumWords.Enabled = True
    Call displayInfo
End Sub

Private Sub optAll_Click()
    txtNumWords.Text = frmGermatik.totNum
    txtNumWords.Enabled = False
    Call displayInfo
End Sub

Private Sub txtNumWords_Change()
    If (Not txtNumWords.Text = "") Then
        If IsNumeric(txtNumWords.Text) Then
            If txtNumWords.Text > frmGermatik.totNum Then
                txtNumWords.Text = frmGermatik.totNum
            End If
        Else
            txtNumWords.Text = frmGermatik.totNum
        End If
    End If
    Call displayInfo
End Sub

Private Sub displayInfo()
    If chkRandom.Value = vbChecked Then
        If optModePractice = True Then
            lblInfo.Caption = "Asks all the words in random order until you get them all right. If you get a word wrong, it will ask you it again."
        Else
        If (optWords = True) And (Not txtNumWords.Text = "") Then
                lblInfo.Caption = "Asks " & txtNumWords & " words selected randomly. If you get a word wrong, it will not be asked again."
        Else
            lblInfo.Caption = "Asks all the words in random order. If you get a word wrong, it will not be asked again."
        End If
        End If
    Else
        If optModePractice = True Then
            lblInfo.Caption = "Asks all the words in sequential order until you get them all right. If you get a word wrong, it will ask you it again."
        Else
            If optWords = True Then
                lblInfo.Caption = "Asks " & txtNumWords & " words in sequential order. If you get a word wrong, it will not be asked again."
            Else
                lblInfo.Caption = "Asks all the words in sequential order. If you get a word wrong, it will not be asked again."
            End If
        End If
    End If
End Sub

Private Sub frameTestModeOptionsDisable()
    optWords.Enabled = False
    txtNumWords.Enabled = False
    optAll.Enabled = False
End Sub

Private Sub frameTestModeOptionsEnable()
    optWords.Enabled = True
    txtNumWords.Enabled = Not optAll
    optAll.Enabled = True
End Sub

Private Sub optAskEnglish_Click()
    If frmGermatik.checkGrammar Then chkGrammar.Value = vbChecked Else chkGrammar.Value = vbUnchecked
    FrameGrammar.Enabled = True
    lblGrammar.Enabled = True
    chkGrammar.Enabled = True
End Sub

Private Sub optAskGerman_Click()
    chkGrammar.Value = vbUnchecked
    FrameGrammar.Enabled = False
    lblGrammar.Enabled = False
    chkGrammar.Enabled = False
End Sub

Private Sub btnClose_Click()
    'update all variables in frmGermatik
    frmGermatik.ModeIsPractice = optModePractice

    If optModePractice Then
        frmGermatik.TestLimit = frmGermatik.totNum
        frmGermatik.optAll = True
    Else
        If optAll Then
            frmGermatik.selectedNumWords = frmGermatik.totNum
        Else
            If txtNumWords.Text = "" Then
                frmGermatik.selectedNumWords = frmGermatik.totNum
            Else
                frmGermatik.selectedNumWords = txtNumWords.Text
            End If
        End If
        frmGermatik.optAll = optAll
    End If

    If chkRandom.Value = vbChecked Then
        frmGermatik.doRandom = True
    Else
        frmGermatik.doRandom = False
    End If
    
    If chkGrammar Then
        frmGermatik.checkGrammar = True
    Else
        frmGermatik.checkGrammar = False
    End If
    
    If optAskEnglish Then
        frmGermatik.askEnglish = True
    Else
        frmGermatik.askEnglish = False
    End If
    
    Unload Me
End Sub

