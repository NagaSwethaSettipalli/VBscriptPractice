Option Explicit

Dim num

num = InputBox("Enter a number between 1 and 25")

Do While num < 25
  WScript.Echo "num : " & num
  
  'if entered num is out of range it should display a message and exit .. this logic is not working recheck and fix
    If num > 25 Then
        MsgBox "number is out of range"
        MsgBox "Exiting!!!!!!"
        Exit Do
    End If
  
     num = num +1
Loop
 