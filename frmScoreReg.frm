VERSION 5.00
Begin VB.Form frmScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "frmScoreReg.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameHighScores 
      Caption         =   "High Scores"
      Height          =   2220
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3495
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   25
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   8
         Left            =   165
         TabIndex        =   24
         Top             =   1920
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   23
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   7
         Left            =   165
         TabIndex        =   22
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   21
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   6
         Left            =   165
         TabIndex        =   20
         Top             =   1440
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   19
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   5
         Left            =   165
         TabIndex        =   18
         Top             =   1200
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   17
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   4
         Left            =   165
         TabIndex        =   16
         Top             =   960
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   3
         Left            =   165
         TabIndex        =   14
         Top             =   720
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   12
         Top             =   480
         Width           =   2130
      End
      Begin VB.Label lblHSscore 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblHSName 
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   10
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame frameScores 
      Caption         =   "Results"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         Caption         =   "40%"
         Height          =   255
         Left            =   2220
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label label4 
         Alignment       =   2  'Center
         Caption         =   "Percent"
         Height          =   255
         Left            =   2220
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "24"
         Height          =   255
         Left            =   540
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCorrect 
         Alignment       =   2  'Center
         Caption         =   "20"
         Height          =   255
         Left            =   1380
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Total"
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Correct"
         Height          =   255
         Left            =   1380
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "670/800"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3315
      Width           =   3495
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public viewOnly As Boolean
Dim currentNumberOfScores As Integer
Dim scorePoints As Integer
Public newName As String

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
    
    'get current high scores from registry
    Call getHighScores
    
    If viewOnly Then
        frameScores.Visible = False
        btnOK.Top = 2455
        Me.Height = 3300
    Else
        Call doNewScore
    End If
End Sub

'this is in a sub so that we can load the form
' to look at it without adding a name
Private Sub doNewScore()
Dim tmpName As String
Dim tmpScore As String
Dim i As Integer
Dim totalPossible As String

    scorePoints = frmGermatik.scoFinalScore
'    totalPossible = (Int(2 * frmGermatik.totNum) + Int(frmGermatik.totNum / 3)) * 10
    'total possible should be based on total number of
    ' words asked
    totalPossible = (Int(2 * frmGermatik.TestLimit) + Int(frmGermatik.TestLimit / 3)) * 10
    
    lblScore = scorePoints & "/" & totalPossible
    lblCorrect = frmGermatik.score 'number right (also number of words in list in practice mode)
    lblTotal = frmGermatik.total   'total ASKED (>= number of words in list)
    lblPercent = Int(100 * frmGermatik.score / frmGermatik.total)
    
    'if there are less than 8 scores, automatically add
    If currentNumberOfScores < 8 Then
        Call addName
        'since we added a score, increase currentNumberOfScores
        currentNumberOfScores = currentNumberOfScores + 1
        Call saveHighScores
    Else
        'if new score is higher, then add to list
        If scorePoints > getScoreFromPercent(lblHSscore.Item(7)) Then
            Call addName
            Call saveHighScores
        End If
    End If

End Sub

Private Sub addName()
'adds a name and score to the lists (made by
' labels). Assumes score is higher than any on
' list, and finds appropriate place for name
' and score.
Dim pos As Integer
Dim i As Integer
Dim place As Integer
    
    ' ask for a name
    frmScoreName.Show vbModal
    
    ' determine where to add name
    place = 1
    Do Until scorePoints > getScoreFromPercent(lblHSscore(place))
        place = place + 1
    Loop
    
    'insert name and score in labels
    ' bump down names
    For i = currentNumberOfScores + 1 To place + 1 Step -1
        If Not i = 9 Then
            lblHSName(i) = lblHSName(i - 1)
            lblHSscore(i) = lblHSscore(i - 1)
        End If
    Next
    ' add name in place
    lblHSName(place) = newName
    lblHSscore(place) = lblScore
    
    ' make new score bold
    lblHSName(place).FontBold = True
    lblHSscore(place).FontBold = True
    
End Sub

Private Sub getHighScores()
Dim i As Integer
Dim fromReg As String
Dim score(9) As String
score(1) = "s1"
score(2) = "s2"
score(3) = "s3"
score(4) = "s4"
score(5) = "s5"
score(6) = "s6"
score(7) = "s7"
score(8) = "s8"
    
    currentNumberOfScores = 0

    For i = 1 To 8
        fromReg = GetSettingString(HKEY_LOCAL_MACHINE, "Software\BlueEyeStudios\Germatik", score(i), "")
        If Not fromReg = "" Then
            currentNumberOfScores = currentNumberOfScores + 1
            lblHSName(currentNumberOfScores) = Left(fromReg, InStr(1, fromReg, ",") - 1)
            lblHSscore(currentNumberOfScores) = Right(fromReg, Len(fromReg) - InStr(1, fromReg, ","))
        End If
    Next

End Sub

Private Sub saveHighScores()
Dim tmpToSave As String
Dim score(9) As String
score(1) = "s1"
score(2) = "s2"
score(3) = "s3"
score(4) = "s4"
score(5) = "s5"
score(6) = "s6"
score(7) = "s7"
score(8) = "s8"
    
    'save modified list to file
    For i = 1 To currentNumberOfScores
        tmpToSave = lblHSName(i) & "," & lblHSscore(i)
        If tmpToSave = "," Then MsgBox i & " - " & lblHSName(i - 1)
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\BlueEyeStudios\Germatik", score(i), tmpToSave
    Next

End Sub

Private Function getScoreFromPercent(score As String) As Integer
    If score = "" Then
        getScoreFromPercent = 0
    Else
        getScoreFromPercent = Val(Left(score, InStr(1, score, "/") - 1))
    End If
End Function
