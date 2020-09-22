VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3840
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2650.436
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "System Memory"
      Height          =   1335
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
      Begin VB.CommandButton btnRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         ToolTipText     =   "Refresh the memory display"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblAvailPhys 
         Caption         =   "0000 kb"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblTotalPhys 
         Caption         =   "0000 kb"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Available"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1920
      Left            =   240
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   1348.48
      ScaleMode       =   0  'User
      ScaleWidth      =   1348.48
      TabIndex        =   1
      Top             =   240
      Width           =   1920
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   1214
   End
   Begin VB.Label lblWeb 
      Caption         =   "http://www.aaronparecki.com"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3035
      Width           =   3135
   End
   Begin VB.Label lblEmail 
      Caption         =   "e-mail: germatik@aaronparecki.com"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2709
      Width           =   3135
   End
   Begin VB.Label lblCopyright 
      Caption         =   "©2002 by Aaron Parecki. All rights reserved."
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5338.509
      Y1              =   1563.343
      Y2              =   1563.343
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmAbout.frx":444E
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   2250
      TabIndex        =   2
      Top             =   960
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      Caption         =   "Germatik"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5324.423
      Y1              =   1573.696
      Y2              =   1573.696
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   2340
      TabIndex        =   5
      Top             =   600
      Width           =   3315
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Programmer: Aaron Parecki"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   2400
      Width           =   3150
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MS As MEMORYSTATUS

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
    Me.Icon = frmGermatik.Icon
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
    lblTitle.Caption = App.Title
    Call dispMem
End Sub

Private Sub btnRefresh_Click()
    Call dispMem
End Sub

Private Sub dispMem()
Dim totalPhys As String
Dim availPhys As String
    
    MS.dwLength = Len(MS)
    GlobalMemoryStatus MS

    If MS.dwTotalPhys / 1024 > 18000 Then
        totalPhys = Round(CDbl(MS.dwTotalPhys) / 1048576, 2) & "mb"
        If MS.dwAvailPhys / 1024 > 40000 Then
            availPhys = Round(CDbl(MS.dwAvailPhys) / 1048576, 2) & "mb"
        Else
            availPhys = MS.dwAvailPhys / 1024 & "kb"
        End If
    Else
        totalPhys = CDbl(MS.dwTotalPhys) / 1024 & "kb"
        availPhys = MS.dwAvailPhys / 1024 & "kb"
    End If
    
    lblTotalPhys.Caption = totalPhys
    lblAvailPhys.Caption = availPhys
    
    'MS.dwMemoryLoad contains percentage memory used
    'MS.dwTotalPhys contains total amount of physical memory in bytes
    'MS.dwAvailPhys contains available physical memory
'    'MS.dwTotalPageFile contains total amount of memory in the page file
'    lblTotPg.Caption = MS.dwTotalPageFile / 1024 & "kb"
'    'MS.dwAvailPageFile contains available amount of memory in the page file
'    lblFreePg.Caption = MS.dwAvailPageFile / 1024 & "kb"
    
    If MS.dwAvailPhys < MS.dwTotalPhys / 4 Then
        lblAvailPhys.ForeColor = &HFF&
    End If
End Sub

