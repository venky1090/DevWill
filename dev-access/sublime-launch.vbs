' Launch sublime
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim strCurDir, filesys , parentDir, programPath
strCurDir = WshShell.CurrentDirectory

'File System Object
Set filesys = CreateObject("Scripting.FileSystemObject") 

' Get parent Dir
parentDir = filesys.GetParentFolderName(strCurDir)

'Chaging working directory to parent
WshShell.CurrentDirectory = parentDir
programPath = WshShell.CurrentDirectory & "\programs\sublime\sublime_text.exe"

WshShell.Run programPath

Set WshShell = Nothing

