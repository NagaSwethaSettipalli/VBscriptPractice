'Built-in String Functions:
'Mid, Len, StrReverse, LCase, UCase, Replace, Left, Right, LTrim, RTrim, Trim, StrComp, InStr

site = "www.CodingGears.com"
message1 = "I am learning VBScripting at CodingGears.com"
message2 = "    CodingGears    "

'Mid: Extracts substring out of a string
result1 = Mid(site, 5, 6) ' you have to provide starting point of string that you want to extract and the no of characters in that string
'result1 = Mid(site, 5)
'DisplayMessage result1, 101

'Len - finds the length of string
result2 = Len(site)
'DisplayMessage result2, 101

'StrReverse - to reverse a string
result3 = StrReverse(site)
'DisplayMessage result3, 101

'LCase - converts everything to lower case
result4 = LCase(site)
'DisplayMessage result4, 101

'UCase - converts everything to upper case
result5 = UCase(site)
'DisplayMessage result5, 101

'Replace - used to replace part of string with something
result6 = Replace(site, "CodingGears", "GlobalTraining")
'DisplayMessage result6, 101

'Left - you can tell how many characters you want to print from left hand side
result7 = Left(site, 3)
'DisplayMessage result7, 101

'Right - you can tell how many characters you want to print from right hand side
result8 = Right(site, 3)
'DisplayMessage result8, 101

'LTrim - removes white spaces on left hand side
result9 = LTrim(message2)
'DisplayMessage Len(result9), 101 'it will print length of message after trimming white spaces on left hand side

'RTrim - removes white spaces on right hand side
result10 = RTrim(message2)
'DisplayMessage Len(result10), 101 'it will print length of message after trimming white spaces on right hand side

'Trim - used to replace part of string with something
result11 = Trim(message2)
'DisplayMessage Len(result11), 101 ' it will print length of message after trimming both sides (left and right)


'StrComp - used to Compare two strings , returns 0 if they match ; if string1 is less than string2 it returns -1; if string1 is greater than string2 returns 1
'also we can use vbtextcompare
'result12 = StrComp("CodingGears", "CodingGears")
result12 = StrComp("CodingGears", "CodingGears", vbTextCompare)
'DisplayMessage result12, 101

'InStr - if you want to know the position of certain character with in a string use this
result13 = InStr(message1, "am" ) 'in  message1 string i want to know position of 'am'
result14 = InStr(message1 , "VBScripting")' this will tell you at what position VBScripting starts
DisplayMessage result13, 101
DisplayMessage result14, 101

Function DisplayMessage(message, id)
    MsgBox id & " : " & message,0,">>>>Welcome<<<<"
End Function