Attribute VB_Name = "modFileAssociations"
Public Type menuCommands
 Captions As New Collection
 Commands As New Collection
End Type

Public Type fileType
 Commands As menuCommands
 Extension As String
 ProperName As String
 FullName As String
 ContentType As String
 IconPath As String
 IconIndex As Integer
End Type

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal _
  hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal _
  hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As _
  Long, lpbData As Any, ByVal cbData As Long) As Long

Public Sub CreateExtension(newfiletype As fileType)

Dim IconString As String
Dim Result As Long, Result2 As Long, ResultX As Long
Dim ReturnValue As Long, HKeyX As Long
Dim cmdloop As Integer

IconString = newfiletype.IconPath & "," & newfiletype.IconIndex

    If Left$(newfiletype.Extension, 1) <> "." Then _
      newfiletype.Extension = "." & newfiletype.Extension
    
    RegCreateKey HKEY_CLASSES_ROOT, newfiletype.Extension, Result
    ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, _
      ByVal newfiletype.ProperName, LenB(StrConv(newfiletype.ProperName, vbFromUnicode)))
    
    'Set up content type
    If newfiletype.ContentType <> "" Then
    ReturnValue = RegSetValueEx(Result, "Content Type", 0, REG_SZ, ByVal _
      CStr(newfiletype.ContentType), LenB(StrConv(newfiletype.ContentType, vbFromUnicode)))
    End If
    
    RegCreateKey HKEY_CLASSES_ROOT, newfiletype.ProperName, Result
    
    If Not IconString = ",0" Then
        RegCreateKey Result, "DefaultIcon", _
            Result2 'Create The Key of "ProperNameDefaultIcon"
        ReturnValue = RegSetValueEx(Result2, _
            "", 0, REG_SZ, ByVal IconString, _
            LenB(StrConv(IconString, vbFromUnicode)))
        'Set The Default Value for the Key
    End If
    
    ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, ByVal newfiletype.FullName, _
        LenB(StrConv(newfiletype.FullName, vbFromUnicode)))
    RegCreateKey Result, ByVal "Shell", ResultX
    
    'Create neccessary subkeys for each command
    For cmdloop = 1 To newfiletype.Commands.Captions.Count
        RegCreateKey ResultX, ByVal newfiletype.Commands.Captions(cmdloop), Result
        RegCreateKey Result, ByVal "Command", Result2
        Dim CurrentCommand$
        CurrentCommand = newfiletype.Commands.Commands(cmdloop)
        ReturnValue = RegSetValueEx(Result2, "", 0, REG_SZ, ByVal CurrentCommand$, _
            LenB(StrConv(CurrentCommand$, vbFromUnicode)))
        RegCloseKey Result
        RegCloseKey Result2
    Next
    
    RegCloseKey Result2
End Sub

