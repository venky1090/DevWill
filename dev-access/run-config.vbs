' set initial config
Set WshShell = WScript.CreateObject("WScript.Shell")

' Set env variable
Dim strCurDir, filesys , xmlObj, xmlPath, nodejsPath, gitPath, finalPath, existingPath, fullPath

' XML object for reading xml
xmlPath = WshShell.CurrentDirectory & "\config.xml"
Set xmlObj = CreateObject("Microsoft.XMLDOM")
xmlObj.async = false

'Read the Xml 
IF xmlObj.load(xmlPath) Then
	WScript.Echo "Loaded settings."
Else
	WScript.Echo "Failed to load settings."
	Set xPE = xmlObj.parseError
	WScript.Echo xPE.reason
	WScript.Quit
End IF

' XML operations
Set root = xmlObj.documentElement
Set devwillTag = root.getElementsByTagName("devwill_path")(0)
Dim devwillPath
devwillPath = devwillTag.text

' define programs path
nodejsPath = devwillPath + "\programs\nodejs;"
gitPath = devwillPath + "\programs\GitPortable\App\Git\bin;"

' Setting programs path
Dim WshSySEnv
Set WshSysEnv = WshShell.Environment("USER")

' Setting DevWill path
WshSysEnv("DEVWILL_HOME") = devwillPath

' Set empty path for misc scenarios
existingPath = WshSysEnv("path")
finalPath = existingPath + ";"
WshSysEnv("path") = finalPath

' nodejs
existingPath = WshSysEnv("path")
finalPath = existingPath + nodejsPath
WshSysEnv("path") = finalPath

' git
existingPath = WshSysEnv("path")
finalPath = existingPath + gitPath
WshSysEnv("path") = finalPath

WScript.Echo "Path set successfully!"

Set WshSySEnv = Nothing



