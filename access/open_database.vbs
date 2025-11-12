' VBScript to open Access database (CarRental.accdb) in the same folder as the script
Option Explicit
Dim fso, scriptPath, dbPath, sh
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = scriptPath & "\CarRental.accdb"
If Not fso.FileExists(dbPath) Then
  WScript.Echo "Database file not found: " & dbPath
  WScript.Quit 1
End If
Set sh = CreateObject("WScript.Shell")
sh.Run """ & dbPath & """, 1, False