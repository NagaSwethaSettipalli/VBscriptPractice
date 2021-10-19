Option Explicit

Dim num1, num2, total, sum

num1 = 10
num2 = 20

total = Add(num1, num2)

WScript.Echo "The sum of two numbers :" & total

'A Function Procedure  -- can return a value
Function Add(num1, num2)
    sum = num1 + num2
    Add = sum
End Function