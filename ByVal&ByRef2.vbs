Option Explicit

Dim mynum1, mynum2, result

mynum1 = 5
mynum2 = 7

result = AddNumbers(mynum1, mynum2)

MsgBox "M1 : Result is " & result ' prints result as 12
MsgBox "M2: mynum1 = " & mynum1 & "and mynum2 = " & mynum2  'prints mynum1 as 1111 mynum2 as 2222


Function AddNumbers(num1, num2) ' if you don't specify explicitly by default it will consider it as ByRef
    AddNumbers = num1 + num2
    num1 = 1111
    num2 = 2222
End Function