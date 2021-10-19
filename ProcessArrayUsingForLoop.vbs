Option Explicit

Dim i, total, arrNums

arrNums = Array(1,2,3,4,5,6)

For i = UBound(arrNums) To LBound(arrNums) Step -1
    total = total + arrNums(i)
    If total > 6 Then
      WScript.Echo " Total is greater than 6 " & vbCrLf & "Current total is : " & total
    End If
Next

WScript.Echo " The sum of all emlements in array is : " & total


