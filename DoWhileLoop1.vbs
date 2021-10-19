Dim a

a = 1

Do While a < 20

        MsgBox "a : " & a
        MsgBox " welcome "
        If a = 5 Then
            MsgBox "a is equal to 5"
            MsgBox "Exiting!!!!!!"
            Exit Do
        End If
      a = a + 1

Loop

MsgBox a & "I'm ouside the loop"