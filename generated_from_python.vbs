MsgBox "Hello, World!"
Dim ScriptFullName
ScriptFullName = WScript.ScriptFullName
ScriptName = WScript.ScriptName
Dim CurrentScriptDirectory
CurrentScriptDirectory = Mid(ScriptFullName, 1, Len(ScriptFullName) - Len(ScriptName))
MsgBox "ScriptFullName = " + ScriptFullName
MsgBox "ScriptName = " + ScriptName
MsgBox "CurrentScriptDirectory = " + CurrentScriptDirectory
