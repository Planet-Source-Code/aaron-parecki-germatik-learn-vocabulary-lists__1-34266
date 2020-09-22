Attribute VB_Name = "modColours"
'This is not my code. Thanks to whoever it was on PSC
'for it. I'm sorry I don't remember who you are.
Option Explicit

Private Const BITSPIXEL = 12
Private Const PLANES = 14

Private Declare Function CreateIC Lib "gdi32" _
    Alias "CreateICA" (ByVal lpDriverName As String, _
    ByVal lpDeviceName As Any, ByVal lpOutput As Any, _
    ByVal lpInitData As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" _
    (ByVal hdc As Long) As Long

Public Function getColorDepth() As Integer
Dim lHdc As Long, lPlanes As Long, lBitsPerPixel As Integer
' Declare variables

lHdc = CreateIC("DISPLAY", 0&, 0&, 0&)
' Create the device context for the display

If lHdc = 0 Then
  ' An error has occurred and the function will exit
  getColorDepth = -1
  Exit Function
End If

lPlanes = GetDeviceCaps(lHdc, PLANES)
' Return info on number of planes

lBitsPerPixel = GetDeviceCaps(lHdc, BITSPIXEL)
' Use display device context to return info on the
' number of pixels

lHdc = DeleteDC(lHdc)
' Delete the device context

Select Case lPlanes

  Case 1
    ' 1 plane is available. This will be the same for most
    ' computer systems

    getColorDepth = lBitsPerPixel
    'returns color depth in bits

'    Select Case lBitsPerPixel
'        ' Select the number of colours available
'      Case 4: GetAvailableColours = "4 Bit, 16 Colours"
'      Case 8: GetAvailableColours = "8 Bit, 256 Colours"
'      Case 16: GetAvailableColours = "16 Bit, 65536 Colours"
'      Case 24: GetAvailableColours = "24 Bit True Colour, 16.7 Million Colours"
'      Case 32: GetAvailableColours = "32 Bit True Colour, 16.7 Million Colours"
'    End Select

  Case 4
    getColorDepth = 16
'    GetAvailableColours = "16 Bit, 65536 Colours"
    ' If there are 4 planes then the availible
    ' colours will be 16-bit

  Case Else
    getColorDepth = -2
    ' The number of colours could not be determined

End Select

End Function


