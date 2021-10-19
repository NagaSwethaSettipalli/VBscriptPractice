Option Explicit

Dim arrElements, item 

arrElements = Array(1,2,3,4,5,6,7,8,9,10)

For Each item in arrElements
    WScript.Echo "Array element is : " & item
Next

