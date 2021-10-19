'syntax : For Each item in [something] 
           'Do something
           'Next
Option Explicit

Dim arrNames, name

arrNames = Array("Tilak", "Swetha","Medha","Megha")

For Each name in arrNames
    WScript.Echo  "name is : " & name 
Next
