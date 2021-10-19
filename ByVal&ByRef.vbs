'using ByRef & ByVal with parameters/ arguments
'ByVal = passed by value
'ByRef = passed by value or reference
'The argument is passed by value if it is enclosed in parantheses
'and if the parantheses do not apply to the parameter list
'when you use call keyword to call a function, you have to enclose the parameter/arguments in parantheses,
'The argument is also passed by value if the variable sent as an argument is in a class.otherwise, it is passed by reference
'If you don't specify explicitly if it is ByVal or ByRef it will treat as ByRef
'when you try to use ByRef on the properties or variables that are part of a class then it is always treated as ByVal


Option Explicit

Dim MyNumber

MyNumber = 100

'PlayWithNumbers MyNumber ' prints 25

'PlayWithNumbers (MyNumber) ' prints 100

'Call PlayWithNumbers (MyNumber) ' prints 25

'Call PlayWithNumbers ((MyNumber)) ' prints 100

'PlayWithNumbers2 MyNumber ' prints 100

'PlayWithNumbers2 (MyNumber) ' prints 100

'Call PlayWithNumbers2 (MyNumber) ' prints 100

'Call PlayWithNumbers2 ((MyNumber)) ' prints 100

MsgBox "M1 : " & MyNumber


Function PlayWithNumbers(ByRef MyParam)
    MyParam = 25
End Function

Function PlayWithNumbers2(ByVal MyParam)
   MyParam = 25
End Function
    
