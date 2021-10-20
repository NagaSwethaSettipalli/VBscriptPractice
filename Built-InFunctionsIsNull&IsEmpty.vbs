'Empty keyword is used to indicate an uninitialized variable value
'Empty means variable that is un initialized , Null means variable is there but the value is null means nothing
Dim var, var_msg, var_null, var_empty

var_data = 100
var_null = Null
var_empty = Empty
var4  = Null

result1 = IsEmpty(var_data)
result2 = IsNull(var_data)

DisplayMessage result1, "IsEmpty"
DisplayMessage result2, "IsNull"

Function DisplayMessage(message, id)
    MsgBox id & " : " & message,0,"Welcome"
End Function