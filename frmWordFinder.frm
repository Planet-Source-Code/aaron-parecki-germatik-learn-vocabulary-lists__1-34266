VERSION 5.00
Begin VB.Form frmWordFinder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Finder"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmWordFinder.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstEnglish 
      Height          =   1815
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox lstGerman 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton btnEnglish 
      Caption         =   "Find From &English"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton btnGerman 
      Caption         =   "Find From &German"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtEnglishIn 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtGermanIn 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblEngSearch 
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblGerSearch 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "English"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "German"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmWordFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colGerman As New Collection
Dim colEnglish As New Collection

Private Sub btnGerman_Click()
Dim pos As Integer
    
    If Not txtGermanIn.Text = "" Then
        lstGerman.clear
        lstEnglish.clear
        
        For pos = 1 To colGerman.Count
            If InStr(1, colGerman.Item(pos), txtGermanIn.Text, 1) Then
                lstGerman.AddItem colGerman.Item(pos)
                lstEnglish.AddItem colEnglish.Item(pos)
            End If
        Next
            
        If lstGerman.ListCount = 0 Then
            lstGerman.AddItem ("No items found")
            lstEnglish.AddItem ("No items found")
        End If

        lblGerSearch.Caption = txtGermanIn.Text
        lblEngSearch.Caption = ""
        txtGermanIn.Text = ""
        txtEnglishIn.Text = ""

    End If
End Sub

Private Sub btnEnglish_Click()
Dim pos As Integer
    
    If Not txtEnglishIn.Text = "" Then
        lstGerman.clear
        lstEnglish.clear
        
        For pos = 1 To colGerman.Count
            If InStr(1, colEnglish.Item(pos), txtEnglishIn.Text, 1) Then
                lstGerman.AddItem colGerman.Item(pos)
                lstEnglish.AddItem colEnglish.Item(pos)
            End If
        Next
            
        If lstGerman.ListCount = 0 Then
            lstGerman.AddItem ("No items found")
            lstEnglish.AddItem ("No items found")
        End If

        lblGerSearch.Caption = ""
        lblEngSearch.Caption = txtEnglishIn.Text
        txtGermanIn.Text = ""
        txtEnglishIn.Text = ""

    End If
End Sub

Private Sub txtGermanIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call btnGerman_Click
    Call specialChars(txtGermanIn, KeyCode, Shift)
End Sub

Private Sub txtEnglishIn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call btnEnglish_Click
End Sub

Private Sub doOpen()
    If Not frmGermatik.CommonDialogList.fileName = "" Then
        If Not doesFileExist(frmGermatik.CommonDialogList.fileName) Then Exit Sub
        Call collClear(colGerman)
        Call collClear(colEnglish)
        Open frmGermatik.CommonDialogList.fileName For Input As #1
            Do While Not EOF(1)
                Input #1, ger
                Input #1, eng
                colGerman.Add (ger)
                colEnglish.Add (eng)
            Loop
        Close #1
        Call displayTitleWithFileName(Me, "Word Finder - ", frmGermatik.CommonDialogList.fileName, False)
    Else
        MsgBox "If you are seeing this message box, then please tell Aaron exactly what you did to get to it.", vbCritical + vbOKOnly, "You should never see this"
    End If
End Sub

''''''''''''''''''''''''''''
' Ties list boxes together '
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
Private Sub lstGerman_Click() 'for arrow keys
    Call lstGerman_MouseDown(0, 0, 0, 0)
End Sub
Private Sub lstEnglish_Click()
    Call lstEnglish_MouseDown(0, 0, 0, 0)
End Sub
'                          '
''''''''''''''''''''''''''''

Private Sub Form_Resize()
Dim tmp As Integer

    Me.Width = 4575
    If Me.Height >= 3500 Then
        lstGerman.Height = Me.Height - 2325
        lstEnglish.Height = lstGerman.Height
        txtGermanIn.Top = Me.Height - 1380
        txtEnglishIn.Top = txtGermanIn.Top
        btnGerman.Top = Me.Height - 900
        btnEnglish.Top = btnGerman.Top
    Else
        Me.Height = 3500
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmListEditor.Icon
    Call doOpen
End Sub

