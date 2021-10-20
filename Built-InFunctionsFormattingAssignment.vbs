Option Explicit

Dim MyNum, EngMarks, SciMarks, MatMarks
Dim total, average , mypercentage

EngMarks = 90
SciMarks = 95
MatMarks = 96

'1. number is 9.3333333333 .. round the value to 3 decimals

MyNum = 9.3333333333
'MsgBox "the value after rounding to 3 decimals is : " & FormatNumber(MyNum, 3)

'2. Adam scored English-90, science -95 Mathematics - 96. Calculate average marks and display it as a message with 3 decimals in % (eg : 80.123%)
 total = EngMarks + SciMarks + MatMarks
 average = total / 3
 mypercentage = FormatPercent(average/100 , 3)

 MsgBox "Total is  :" & total
 MsgBox "Average is  :" & average
 MsgBox "Total percentage of marks Amit scored is :" & mypercentage

