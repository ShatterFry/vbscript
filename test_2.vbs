CurrentScriptDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "Current script directory: " & CurrentScriptDirectory