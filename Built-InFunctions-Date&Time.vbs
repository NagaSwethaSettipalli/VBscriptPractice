'Using built-in functions - working with dates
'Now, Date, DatePart, DateAdd, Weekday, Day, Month, MonthName, Year, Weekday, WeekdayName
'DateAdd, DateDiff, DataValue
'Time, Hour, Minute, Second, TimeValue

mydate_date = Date 'Date just gives you the date with out the time stamp
mydate_now = Now  'Now returns current date and time stamp
mydate_set = #03-15-2015# ' How to set a desired date
month1 = DatePart("m",mydate_set) 'DatePart - Extracts month/day/year out of particular date
day1 = DatePart("d",mydate_set)
year1 = DatePart("yyyy", mydate_set)
month1_name = MonthName(DatePart("m",mydate_set)) 'If you want to extract name of month you can use MonthName()
mynewdate1 = DateAdd("d", 2, mydate_set ) ' d represents days 2 represents 2 days and we are adding 2 days to setdate
mynewdate2 = DateAdd("m", 2, mydate_set )' If you want to add 2 months to set date do this

mytime = Time 'Returns current time
myhour = Hour(mytime) 'To extract Hours from the time
myminute = Minute(mytime) 'To extract minutes from the time
mysecond = Second(mytime) 'To extract seconds from the time
weekday1 = Weekday(date) 'To dispaly which day of week it is in integer format 
weekday2 = WeekdayName(weekday1) 'To display which day of week it is in string format
weekday3 = WeekdayName(Weekday(Now)) 'Instead of using above 2 statements you can using this everything in one line
check_date1 = IsDate(mydate_set) 'Checks if it is date or not and ensures if it is in right date format or not


'Now - returns current date and time stamp
'DisplayMessage mydate_now, "mydate_now"
'*********************************************
'Date- If you want to just get the date with out the time stamp
'DisplayMessage mydate_date, "mydate"
'**********************************************
'DisplayMessage mydate_set, "mydate_set"
'**********************************************
'DatePart - Extracting month/day/year out of particular date
'DisplayMessage month1, "month1"
'DisplayMessage day1, "day1"
'DisplayMessage year1, "year1"
'***********************************************
'DisplayMessage month1_name, "month1_name" 'If you want to extract name of month you can do thisuse MonthName()
'***********************************************
'DateAdd - Add Dates to existing date means you have a date and if you add 10 days to that date what would be the date after 10 days
'DisplayMessage mynewdate1, "mynewdate1"
'DisplayMessage mynewdate2, "mynewdate2"
'***********************************************
'Time - Time function returns current time
'DisplayMessage mytime, "mytime"
'*********************************To extract Hours from the time**************
'Hour - To extract Hours from the time
'DisplayMessage myhour, "myhour"
'***********************************************
'Minute - To extract minutes from the time
'DisplayMessage myminute, "myminute"
'***********************************************
'Second - To extract seconds from the time
'DisplayMessage mysecond, "mysecond"
'***********************************************
'Weekday - displays which day it is
'for eg : If it is Friday and time is 5:00 clock then you need to trigger something some archival work we can use this weekday
'DisplayMessage weekday1, "weekday1"
'DisplayMessage weekday2, "weekday2"
'DisplayMessage weekday3, "weekday3"
'***********************************************
'IsDate - Checks if it is date or not and ensures if it is in right date format or not
DisplayMessage check_date1, "check_date1"

Function DisplayMessage(message, id)
    MsgBox id & " : " & message, 0, "Welcome"
End Function