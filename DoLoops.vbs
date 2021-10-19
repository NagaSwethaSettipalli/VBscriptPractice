
Dim count

count = 1

DO

MsgBox count & " I'm inside the loop"
count = count +1

If count = 5 Then
    ' WScript.Quit - kills the whole loop that's why Im ouside the loop not printed
Exit Do ' Instead of wscript.Quit use EXit DO
    
End If

Loop

MsgBox count & " I'm outside the loop"