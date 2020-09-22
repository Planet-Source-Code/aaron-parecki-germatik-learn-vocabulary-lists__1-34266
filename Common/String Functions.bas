Attribute VB_Name = "modStringFunctions"
'This is my original code

Public Function isLtrCapital(word As String, ltrNum As Integer) As Boolean
''''''''''''''''''''''''''''''''''''''''''
' Determines whether or not the ltrNumTH '
' letter of word is capitalized          '
''''''''''''''''''''''''''''''''''''''''''
    If Mid(word, ltrNum, 1) >= "A" And Mid(word, ltrNum, 1) <= "Z" Then
        isLtrCapital = True
    Else
        isLtrCapital = False
    End If
End Function

Public Function removeStr(str As String, ch As String) As String
''''''''''''''''''''''''''''''''''''''''
' removes character ch from string str '
''''''''''''''''''''''''''''''''''''''''
Dim i As Integer
Dim tempStr(60) As String 'probably big enough to hold a word
Dim newString As String
Dim numChars As Integer

    newString = ""
    numChars = 0
    
    For i = 1 To Len(str)
        If Not (Mid(str, i, 1) = Left(ch, 1)) Then
            tempStr(numChars) = Mid(str, i, 1)
            numChars = numChars + 1
        End If
    Next i

    For i = 1 To numChars
        newString = newString & tempStr(i - 1)
    Next i

    removeStr = newString
End Function

Public Function lastInStr(str As String, ch As String) As Integer
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
