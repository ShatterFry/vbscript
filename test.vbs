Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = "myShortcut.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)

Dim ScriptFullName
ScriptFullName = WScript.ScriptFullName
ScriptName = WScript.ScriptName
Dim CurrentScriptDirectory
CurrentScriptDirectory = Mid(ScriptFullName, 1, Len(ScriptFullName) - Len(ScriptName))

TargetPath = CurrentScriptDirectory & "test.txt"
WScript.Echo "TargetPath = " & TargetPath
oLink.TargetPath = TargetPath
oLink.Save