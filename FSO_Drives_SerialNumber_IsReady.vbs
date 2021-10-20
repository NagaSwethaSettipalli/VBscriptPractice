'Drive.SerialNumber
'Drive.IsReady

Dim oFSO, oDrive, cDrives
Dim disktype

Set  oFSO = CreateObject("Scripting.FileSystemObject")
Set oDrive = oFSO.GetDrive("C:")

MsgBox "Serial Number : " & oDrive.SerialNumber, 0 ,"Serial Number (C: ) :"
MsgBox "Is the drive Ready : " & oDrive.IsReady, 0 ,"Drive Information :"