' Adam wants to fly to france today and would like to return back after 10 days.His flight takes off exactly at 06:30:45 PM
' Display Adam'take-off date
' Dispaly Adam's return date

Option Explicit
Dim takeoffdate, returndate
takeoffdate = Date 
returndate = DateAdd("d", 10, takeoffdate )

MsgBox "Adam's flight take off at : " & takeoffdate
MsgBox "Adam's flight returns at : " & returndate