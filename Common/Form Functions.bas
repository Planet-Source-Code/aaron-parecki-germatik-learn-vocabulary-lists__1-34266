Attribute VB_Name = "modFormFunctions"
'This is my original code except for getX and getY

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Public Sub displayTitleWithFileName(frm As Form, str As String, fileName As String, dispExt As Boolean)
''''''''''''''''''''''''''''''''''''''''''''
' displays str plus the file name part of  '
' string fileName on the title of form frm '
''''''''''''''''''''''''''''''''''''''''''''
Dim fileNameOnly As String
    
    fileNameOnly = Right(fileName, Len(fileName) - lastInStr(fileName, "\"))
    If Not dispExt Then
        fileNameOnly = Left(fileNameOnly, InStr(1, fileNameOnly, ".") - 1)
    End If
    frm.Caption = str & "'" & fileNameOnly & "'"

End Sub

Public Sub centerForm(frm As Form)
    frm.Left = (getX * 15 / 2) - (frm.Width / 2)
    frm.Top = (getY * 15 / 2) - (frm.Height / 2)
End Sub


Private Function getX() As Integer
    getX = GetSystemMetrics(SM_CXSCREEN)
End Function

Private Function getY() As Integer
    getY = GetSystemMetrics(SM_CYSCREEN)
End Function

