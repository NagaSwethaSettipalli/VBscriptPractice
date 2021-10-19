Option Explicit

Dim a

For a = 0 To 5
    WScript.Echo a & " : VBScript learning!!!"
Next

WScript.Echo "-----------------"

For a = 0 to 10 step 2 ' by doing step 2 we can increment value of a by 2 instead of 1
    WScript.Echo a & " : VBScripting is fun...."
Next

WScript.Echo "*****************"

For a = 50 to 10 step -10  ' for doing it in reverse action
    WScript.Echo a & " : VBS#####"
Next