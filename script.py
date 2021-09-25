import subprocess

vscript_file_name = "generated_from_python.vbs"
with open(vscript_file_name, "w") as file_handle:
    #file_handle.write("123\n")
    #file_handle.write("456")

    file_handle.write('MsgBox "Hello, World!"\n')

    file_handle.write('Dim ScriptFullName\n')
    file_handle.write('ScriptFullName = WScript.ScriptFullName\n')

    file_handle.write('ScriptName = WScript.ScriptName\n')

    file_handle.write('Dim CurrentScriptDirectory\n')
    file_handle.write('CurrentScriptDirectory = Mid(ScriptFullName, 1, Len(ScriptFullName) - Len(ScriptName))\n')

    file_handle.write('MsgBox "ScriptFullName = " + ScriptFullName\n')
    file_handle.write('MsgBox "ScriptName = " + ScriptName\n')
    file_handle.write('MsgBox "CurrentScriptDirectory = " + CurrentScriptDirectory\n')

proc = subprocess.run("cscript {}".format(vscript_file_name))
print(proc)