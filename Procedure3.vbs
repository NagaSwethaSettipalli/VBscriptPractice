Option Explicit
Dim num1, num2, result

num1 = 10
num2 = 20

result = Multiply(num1, num2)

DisplayMessage "The result is", result

Sub DisplayMessage(strMessage, intResult)
  MsgBox strMessage & " : " & intResult
End Sub

Function Add(num1, num2)
    Add = num1 +num2
End Function

Function Subtract(num1, num2)
    Subtract = num1 - num2
End Function

Function Multiply(num1, num2)
    Multiply = num1 * num2
End Function

Function Divide(num1, num2)
    Divide = num1 / num2
End Function     