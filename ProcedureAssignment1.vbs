Dim input1, input2, total

input1 = InputBox("Enter the first number: ")
input2 = InputBox("Enter the second number: ")

total = Add(CInt(input1) ,CInt(input2))
MsgBox "The sum of two numbers : " & total


Function Add(num1, num2)
    sum = num1 + num2
    Add = sum
End Function