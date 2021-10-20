number1 = 12345.6789123
number2 = 12568956.256

'Percentage Formatting
'FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)

mypercent1 = FormatPercent(45/80) ' Default
mypercent2 = FormatPercent(45/80, 8, vbTrue) ' 8 decimals, leading 0's
mypercent3 = FormatPercent(-45/80, 8, vbTrue, vbTrue) ' Indiactes whether or not to place negative values with parenthesis, If you set 4th parameter
                                                      ' as vbTrue instead of negative number is will show in paranthesis
mypercent4 = FormatPercent(-45/80, 8, vbTrue, vbTrue, True) ' Indicates whether or not numbers are grouped

'DisplayMessage mypercent4, "MyPercent"

'Currency Formatting 
'FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
 DisplayMessage FormatCurrency(number2, 2), "2 Decimals"
 DisplayMessage FormatCurrency(number1, 12, vbTrue, vbTrue), "12 Decimals, leading 0's, use ()"
 DisplayMessage FormatCurrency(number2, 5, vbTrue, , vbTrue), "5 Decimals, leading 0's, use (), grouping" 'If you want to enable grouping you have to set 4th parameter as vbTrue

Function DisplayMessage(message, id)
    MsgBox id & " : " & message,0,"Welcome"
End Function