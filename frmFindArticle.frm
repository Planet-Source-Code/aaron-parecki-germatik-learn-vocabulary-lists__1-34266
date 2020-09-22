VERSION 5.00
Begin VB.Form frmFindArticle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Article"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnFind 
      Caption         =   "Find Article"
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtNounIn 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblWholeWord 
      Alignment       =   2  'Center
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
      Left            =   150
      TabIndex        =   3
      Top             =   780
      Width           =   2415
   End
   Begin VB.Label lblEnterNoun 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter noun to find its article "
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmFindArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colGerman As New Collection

Private Sub btnFind_Click()
Dim articleIsFound As Boolean
Dim i As Integer
    
    i = 1
    articleIsFound = False
    
    txtNounIn = Trim(txtNounIn)
    
    Do Until articleIsFound Or i > colGerman.Count
        If LCase(Right(colGerman.Item(i), Len(colGerman.Item(i)) - 4)) = LCase(txtNounIn) Then
            If isNoun(colGerman.Item(i)) Then
                lblWholeWord = colGerman.Item(i)
                articleIsFound = True
            End If
        End If
        i = i + 1
    Loop
    
    If Not articleIsFound Then
        lblWholeWord = "Not found"
    End If

End Sub

Private Sub Form_Load()
    Me.Icon = frmListEditor.Icon
    
    If doOpen = False Then
        disableAndDisplay "No list open"
    Else
        If colGerman.Count = 0 Then
            disableAndDisplay "No nouns in list"
        End If
    End If
    
End Sub

Private Sub disableAndDisplay(str As String)
    For Each ctl In Me
        ctl.Enabled = False
    Next
    lblWholeWord = str
End Sub

Private Function doOpen() As Boolean
Dim ger As String
Dim eng As String
Dim filenametoopen As String
    filenametoopen = frmGermatik.CommonDialogList.fileName
    If filenametoopen = "" Then
        doOpen = False
        Exit Function
    End If
        
    If doesFileExist(filenametoopen) Then
        collClear colGerman
        Open filenametoopen For Input As #1
            Do While Not EOF(1)
                Input #1, ger
                Input #1, eng
                If isNoun(ger) Then colGerman.Add ger
            Loop
        Close #1
        doOpen = True
    Else
        MsgBox "File not found: '" & fileNameOnly(filenametoopen) & "'", vbCritical + vbOKOnly, "Error"
        doOpen = False
    End If

End Function

Private Sub txtNounIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call btnFind_Click
        Exit Sub
    End If
    Call specialChars(txtNounIn, KeyCode, Shift)
End Sub
