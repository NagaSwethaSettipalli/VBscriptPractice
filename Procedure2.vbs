'Sub Procedure -- does not return a value
'Function Procedure -- can return a value

Option Explicit

Dim temp, inCelsius, tempInput

Call ConvertTemp

Sub ConvertTemp
    temp = InputBox("Please enter temperature in degrees F.")
    MsgBox "The temperature is " & Round(Celsius(temp), 2) & "degrees Celsius"
End Sub

Function Celsius(fDegrees)
    Celsius = (fDegrees - 32) * 5 / 9 
End Function