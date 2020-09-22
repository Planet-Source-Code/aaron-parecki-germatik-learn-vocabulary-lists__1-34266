Attribute VB_Name = "modMathFunctions"
Function Round(nValue As Double, nDigits As Integer) As Double
'''''''''''''''''''''''''''''''''''''''
' Rounds a number to n decimal places '
'''''''''''''''''''''''''''''''''''''''
    Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)

End Function

Function Remainder(divideThis, byThis) As Integer
'This is my original code
    Remainder = Int(divideThis / byThis) * byThis
    Remainder = divideThis - Remainder
End Function
