' FSO.GetDrive, Drive.AvailableSpace,Drive.FreeSpace,Drive.TotalSize

Option Explicit

Dim oFSO, oDrive, cDrives

Set  oFSO = CreateObject("Scripting.FileSystemObject")

Set oDrive = oFSO.GetDrive("C:")

'the value returned will be in bytes so we need to divide by 1024 3 times to get value in GB
MsgBox "Available space: " &  FormatNumber(((oDrive.AvailableSpace/1024)/1024)/1024, 0) & " GB",0, "C Drive"
MsgBox "Free space: " &  FormatNumber(((oDrive.FreeSpace/1024)/1024)/1024, 0) & " GB",0, "C Drive"
MsgBox "Total size: " &  FormatNumber(((oDrive.TotalSize/1024)/1024)/1024, 0) & " GB",0, "C Drive"