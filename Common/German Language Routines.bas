Attribute VB_Name = "modGerLanguage"
'This is my original code

Public Function specialChars(guess As Object, KeyCode As Integer, Shift As Integer)

    pos = guess.SelStart
    
    If Shift = 5 Then
        ' Caps
        If KeyCode = 65 Then
            guess.Text = Left(guess.Text, pos) + Chr(196) + Right(guess.Text, Len(guess.Text) - pos)  ' A 196
            guess.SelStart = pos + 1
        ElseIf KeyCode = 79 Then
            guess.Text = Left(guess.Text, pos) + Chr(214) + Right(guess.Text, Len(guess.Text) - pos)  ' O
            guess.SelStart = pos + 1
        ElseIf KeyCode = 85 Then
            guess.Text = Left(guess.Text, pos) + Chr(220) + Right(guess.Text, Len(guess.Text) - pos)  ' U
            guess.SelStart = pos + 1
        ElseIf KeyCode = 83 Then
            guess.Text = Left(guess.Text, pos) + Chr(223) + Right(guess.Text, Len(guess.Text) - pos)  ' SS
            guess.SelStart = pos + 1
        End If
    ElseIf Shift = 4 Then
        'lowercase
        If KeyCode = 65 Then
            guess.Text = Left(guess.Text, pos) + Chr(228) + Right(guess.Text, Len(guess.Text) - pos)  ' a
            guess.SelStart = pos + 1
        ElseIf KeyCode = 79 Then
            guess.Text = Left(guess.Text, pos) + Chr(246) + Right(guess.Text, Len(guess.Text) - pos)  ' o
            guess.SelStart = pos + 1
        ElseIf KeyCode = 85 Then
            guess.Text = Left(guess.Text, pos) + Chr(252) + Right(guess.Text, Len(guess.Text) - pos)  ' u
            guess.SelStart = pos + 1
        ElseIf KeyCode = 83 Then
            guess.Text = Left(guess.Text, pos) + Chr(223) + Right(guess.Text, Len(guess.Text) - pos)  ' ss
            guess.SelStart = pos + 1
        End If
    End If

End Function

Public Function isNoun(word As String) As Boolean
''''''''''''''''''''''''''''''''''''''''
' Determines whether word is a german  '
' noun by looking for der, die, or das '
' in the beginning                     '
''''''''''''''''''''''''''''''''''''''''
    If Len(LTrim(RTrim(word))) < 5 Then
        isNoun = False
    Else
        If Left(LTrim(word), 4) = "der " Or Left(LTrim(word), 4) = "die " Or Left(LTrim(word), 4) = "das " Then
            isNoun = True
        Else
            isNoun = False
        End If
    End If
End Function

Public Function gerWordIsInList(lst As ListBox, lstItem As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''
' checks to see if a word appears in a list '
' returns true also if noun is in list with '
' different article                         '
'''''''''''''''''''''''''''''''''''''''''''''
Dim a As Integer

    For a = 0 To lst.ListCount - 1
        If lst.List(a) = lstItem Then
            gerWordIsInList = True
            Exit Function
        End If
        If getOnlyNoun(lst.List(a)) = getOnlyNoun(lstItem) Then
            gerWordIsInList = True
            Exit Function
        End If
    Next

    gerWordIsInList = False
End Function

Public Function getOnlyNoun(noun As String) As String
    If isNoun(noun) Then
        getOnlyNoun = Right(noun, Len(noun) - 4)
    Else
        getOnlyNoun = ""
    End If
End Function
