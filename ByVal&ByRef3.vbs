'when you try to use ByRef on the properties or variables that are part of a class then it is always treated as ByVal
Class Student
    Public MyStudentId
End Class

Sub ChangeStudentId (ByRef MyId)
    MyId = 5555
End Sub

Dim student1
Set student1 = new Student
student1.MyStudentId = 1111

ChangeStudentId student1.MyStudentId
MsgBox "M1 : The updated Student ID is " & student1.MyStudentId 'when you try to use ByRef on the properties or variables that are part of a class then it is always treated as ByVal
