Option Explicit

Dim roles, name

name = InputBox("Enter name here")

Select Case name

   Case "Thompson"
    MsgBox "Thompson is a President"

   Case "Rooter"
    MsgBox "Rooter is a Sr. Vice President"

   Case "Cooper"
    MsgBox "Cooper is a Vice President"

   Case "Parker"
    MsgBox "Parker is a Manager"

   Case Else
    MsgBox " Hello " & name & ", you are not part of the Management Team."

End Select 
