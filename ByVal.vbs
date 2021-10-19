Option Explicit

Dim MyNumber

MyNumber = 100

PlayWithNumbers MyNumber
MsgBox "M1:" & MyNumber ' Prints 100 original value will be printed when you pass by value
PlayWithNumbers (MyNumber)
MsgBox "M2:" & MyNumber ' Prints 100 original value will be printed even if you add parantheses when you passed by value

Function PlayWithNumbers(ByVal MyParam)
    MyParam = 25
End Function
