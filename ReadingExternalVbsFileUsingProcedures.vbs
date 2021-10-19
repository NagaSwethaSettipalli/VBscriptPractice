Option Explicit

'sub to read external vbs file
Sub Include(extVBScriptFile)
    Dim objFso, objExtFile
    Dim strfileContent, strScriptDir

    strScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objExtFile = objFso.OpenTextFile(strScriptDir & "\" & extVBScriptFile, 1)
    strfileContent = objExtFile.ReadAll
    objExtFile.Close
    ExecuteGlobal strfileContent
    Set objFso = Nothing
    Set objExtFile = Nothing
End Sub

Include "Procedure3.vbs"

Dim num1, num2

result = Subtract(15,35)

DisplayMessage "The result is", result