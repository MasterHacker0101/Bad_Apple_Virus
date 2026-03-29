Set shell = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\Bad_Apple.py"
shell.Run "cmd /c pyw """ & scriptPath & """ || pythonw """ & scriptPath & """", 0, False
