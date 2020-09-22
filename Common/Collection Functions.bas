Attribute VB_Name = "modCollectionFunctions"
'This is my original code

Public Function wordIsInColl(col As Collection, lstItem As String, caseSense As Boolean) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''
' checks to see if a word appears in a collection '
' and if so, returns the position, or -1 if not   '
'''''''''''''''''''''''''''''''''''''''''''''''''''

If caseSense Then
    For a = 1 To col.Count
        If col.Item(a) = lstItem Then
            wordIsInColl = a
            Exit Function
        End If
    Next
Else
    For a = 1 To col.Count
        If LCase(col.Item(a)) = LCase(lstItem) Then
            wordIsInColl = a
            Exit Function
        End If
    Next
End If

'    For a = 1 To col.Count
''        If InStr(1, col.Item(a), lstItem, caseSenseInt) Then
'        If LCase(col.Item(a)) = LCase(lstItem) Then
'            wordIsInColl = a
'            Exit Function
'        End If
'    Next

    wordIsInColl = -1
End Function

'Public Function wordPartIsInColl(col As Collection, lstItem As String, startPos As Integer, caseSense As Boolean) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''
' checks to see if a partial word appears in a  '
' collection  and if so, returns the position,  '
' or -1 if not                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''
'
'If startPos > col.Count Then Exit Function
'
'If caseSense Then
'    For a = startPos To col.Count
'        If InStr(1, col.Item(a), lstItem, 1) Then
'            wordPartIsInColl = a
'            Exit Function
'        End If
'    Next
'Else
'    For a = startPos To col.Count
'        If InStr(1, col.Item(a), lstItem, 0) Then
'            wordPartIsInColl = a
'            Exit Function
'        End If
'    Next
'End If
'
'    wordPartIsInColl = -1
'End Function

Public Sub collAddItem(coll As Collection, sItem As String, pos As Integer, caseSense As Boolean)
''''''''''''''''''''''''''''''''''''
' adds an item to a collection     '
' only if it doesn't already exist '
''''''''''''''''''''''''''''''''''''
    
'set pos to -1 if adding to end of list
'otherwise, inserts item at pos

    ' works with collections
    
    If caseSense Then
        For a = 1 To coll.Count
            If coll.Item(a) = sItem Then Exit Sub
        Next
    Else
        For a = 0 To coll.ListCount - 1
            If LCase(coll.Item(a)) = LCase(sItem) Then Exit Sub
        Next
    End If
    
    If pos = -1 Then
        coll.Add sItem
    Else
        coll.Add sItem, pos
    End If

End Sub

Public Sub collClear(coll As Collection)
    For i = 1 To coll.Count
        coll.Remove (1)
    Next
End Sub

