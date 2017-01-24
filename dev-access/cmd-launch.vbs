' Launch cmd prompt with node access
Set WshShell = WScript.CreateObject("WScript.Shell")

' Set env variable

WshShell.Run "C:/WINDOWS/system32/cmd.exe"
WScript.Sleep 1000

WshShell.SendKeys "cd .."
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "cd workspace"
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "cls"
WshShell.SendKeys "{ENTER}"

Set WshShell = Nothing

