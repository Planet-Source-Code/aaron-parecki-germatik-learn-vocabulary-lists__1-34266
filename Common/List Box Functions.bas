Attribute VB_Name = "modListBoxFunctions"
'This is my original code

Public Sub changeList(listName As ListBox, str As String, pos)
''''''''''''''''''''''''''''''''
' changes an item in a listBox '
' set pos to position of word, '
' or -1 to use selected        '
''''''''''''''''''''''''''''''''
    
    If pos = -1 Then pos = listName.ListIndex
    
    listName.RemoveItem pos
    listName.AddItem str, pos
End Sub

Public Sub listAddItem(lst As ListBox, sItem As String, pos As Integer, caseSense As Boolean)
''''''''''''''''''''''''''''''''''''
' adds an item to a listBox        '
' only if it doesn't already exist '
''''''''''''''''''''''''''''''''''''
Dim a As Integer
    
'set pos to -1 if adding to end of list
'otherwise, inserts item at pos

    ' works with List, Combo
    
    If caseSense Then
        For a = 0 To lst.ListCount - 1
            If lst.List(a) = sItem Then Exit Sub
        Next
    Else
        For a = 0 To lst.ListCount - 1
            If LCase(lst.List(a)) = LCase(sItem) Then Exit Sub
        Next
    End If
    
    If pos = -1 Then
        lst.AddItem sItem
    Else
        lst.AddItem sItem, pos
    End If

End Sub

Public Function wordIsInList(lst As ListBox, lstItem As String, caseSense As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''
' checks to see if a word appears in a list '
'''''''''''''''''''''''''''''''''''''''''''''
Dim a As Integer

    If caseSense Then
        For a = 0 To lst.ListCount - 1
            If lst.List(a) = lstItem Then
                wordIsInList = True
                Exit Function
            End If
        Next
    Else
        For a = 0 To lst.ListCount - 1
            If LCase(lst.List(a)) = LCase(lstItem) Then
                wordIsInList = True
                Exit Function
            End If
        Next
    End If

    isWordInList = False
End Function

Public Function listFindPos(lst As ListBox, word As String, caseSense As Boolean) As Integer
'''''''''''''''''''''''''''''''''''''''''''''
' returns the index of the word in the list '
' or -1 if it doesn't appear at all          '
'''''''''''''''''''''''''''''''''''''''''''''
Dim a As Integer

    If caseSense Then
        For a = 0 To lst.ListCount - 1
            If lst.List(a) = word Then
                listFindPos = a
                Exit Function
            End If
        Next
    Else
        For a = 0 To lst.ListCount - 1
            If LCase(lst.List(a)) = LCase(word) Then
                listFindPos = a
                Exit Function
            End If
        Next
    End If
    
    listFindPos = -1
    
End Function
