'Develop a math calculator with functions Add, subtract, Multiply, Divide X to the power of Y using 2 numbers

Option Explicit

Dim num1, num2, result

num1 = InputBox("Enter num1")
num2 = InputBox("Enter num2")

result = Exponent(num1, num2)
DisplayMessage "The result is :" , result


Sub DisplayMessage(strMessage, intResult)
    MsgBox strMessage & " : " & intResult
  End Sub

Function Add(num1, num2)
    Add = num1 + num2
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

Function Exponent(num1, num2)
    Exponent = num1 ^ num2
End Function