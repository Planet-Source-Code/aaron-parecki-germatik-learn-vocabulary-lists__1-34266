Attribute VB_Name = "modFileFunctions"
'This is my original code

Public Function fileNameOnly(fileName As String) As String
    fileName = Replace(fileName, "/", "\")
    fileNameOnly = Right(fileName, Len(fileName) - lastInStr(fileName, "\"))
End Function

Public Function pathOnly(fileName As String) As String
    fileName = Replace(fileName, "/", "\")
    pathOnly = Left(fileName, lastInStr(fileName, "\"))
End Function

Public Function doesFileExist(fileName As String) As Boolean
On Error GoTo errLabel
    Open fileName For Input As #9
    Close #9
    doesFileExist = True
    Exit Function
errLabel:
    doesFileExist = False
End Function

Private Function lastInStr(str As String, ch As String) As Integer
''''''''''''''''''''''''''''''''''''''
' returns position of last occurence '
' of ch in string str, or 0 if none  '
''''''''''''''''''''''''''''''''''''''
Dim pos As Integer
Dim tempPos As Integer
pos = 0
tempPos = 1
    
    Do While Not tempPos = 0
        tempPos = InStr(IIf(tempPos = 1, 1, tempPos + 1), str, Left(ch, 1))
        If Not tempPos = 0 Then pos = tempPos
    Loop

    lastInStr = pos
End Function
