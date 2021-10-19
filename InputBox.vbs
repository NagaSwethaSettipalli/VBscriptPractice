Const site ="Google"
Dim input1, input2, sum, name , total
' getting input from user
'simple
name = InputBox("Enter your First name: ")
'Adding little Customization: default text
input1 = InputBox("Enter the first number: ", site, "Enter input here")
'More customization: Move the input box around
input2 = InputBox("Enter the second number: ", site, "Enter input here", 1000, 5000)

sum = input1 + input2 ' here if you give 2 numbers as 2 and 3 it will print 23
total = CInt(input1) + CInt(input2) ' here if you give 2 and 3 it will print 5 ; if you enter numbers as A and B it will give error "Microsoft VBScript runtime error:: Type mismatch : Cint"

MsgBox "Hello :" & name & " ! ! ! ", 0, site
MsgBox "The sum of 2 numbers : " & sum, 64, site
MsgBox "The total of 2 numbers : " & total, 64, site