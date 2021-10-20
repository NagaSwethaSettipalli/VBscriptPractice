'Number format, date time format, Format percentage, Format currency

number1 = 1234567.8765432
number2 = 12345678.987

'Number Format:
'FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)

'DisplayMessage FormatNumber(number1, 2), "2 decimals" ' FormatNumber functions takes the number and how many digits you want to display after decimal point
'DisplayMessage FormatNumber(number1, 12, vbTrue),"12 decimals, leading 0s" ' it gives you 12 numbers after decimal point leading with zeros
'DisplayMessage FormatNumber(number1, 8, vbTrue, vbFalse),"8 decimals, leading 0s, use()"
'DisplayMessage FormatNumber(number2, 2, vbTrue, vbTrue, vbFalse),"8 decimals,leading 0s, use()"' if you do vbTrue you will see commas,if you set vbFalse you won't see commas in the number

'FormatDateTime:
'FormatDateTime(Date, NamedFormat) --> useful when you want to generate reports or when you want to write to a file
'NamedFormat used to tell which format you want to dispaly the date

dt1 = FormatDateTime(Date, vbLongDate)    ' weekday, monthname, year
dt2 = FormatDateTime(Date, vbShortDate)   ' mm/dd/yyyy
dt3 = FormatDateTime(Date, vbLongTime)    ' hh:mm:ss PM/AM
dt4 = FormatDateTime(Date, vbShortTime)   ' hh:mm
dt5 = FormatDateTime(Date, vbGeneralDate) ' Default mm/dd/yyyy

DisplayMessage dt1, "vbLongDate"
DisplayMessage dt2, "vbShortDate"
DisplayMessage dt3, "vbLongTime"
DisplayMessage dt4, "vbShortTime"
DisplayMessage dt5, "vbGeneralDate"


Function DisplayMessage(message, id)
   MsgBox id & " : " & message,0,"Welcome"
End Function