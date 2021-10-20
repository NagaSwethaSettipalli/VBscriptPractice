'object.Drives Property
  '  object is always FileSystemObject
  '  Returns a collection with drive objects
'object.DriveLetter
  'object is always a Drive Object
  'Returns the driveletter
'object.Count Property
  'object is always a collection
  'Returns total items in a collection

  Option Explicit

  Dim oFSO, oDrive, cDrives
  Dim ListOfDrives

  Set  oFSO = CreateObject("Scripting.FileSystemObject")
  Set cDrives = oFSO.Drives

  For Each oDrive in cDrives
    MsgBox "Drive letter : " & oDrive.DriveLetter, 0, "Drive on your computer: "
    ListOfDrives = ListOfDrives & "  " & oDrive.DriveLetter
  Next

  MsgBox "Number of Drives : " & cDrives.Count, 0,"Drives on your computer: "
  MsgBox "Drive letters : " & ListOfDrives, 0,"Drives on your computer: "
