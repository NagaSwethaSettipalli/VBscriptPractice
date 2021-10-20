'Drive.DriveExists

Option Explicit

Dim  oFSO, drive
drive = "C:\"

Set  oFSO = CreateObject("Scripting.FileSystemObject")

If oFSO.DriveExists(drive)="True" Then
    MsgBox "We found the drive" & drive, 0, "Result"
Else
    MsgBox "We did not find the drive" & drive, 0, "Result"
End  If