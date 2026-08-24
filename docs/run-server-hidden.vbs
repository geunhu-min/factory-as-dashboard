Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
nodePath = "C:\Program Files\nodejs\node.exe"
If Not fso.FileExists(nodePath) Then
  nodePath = "node"
End If
shell.Run "cmd /c cd /d """ & folder & """ && """ & nodePath & """ server.js > server.log 2>&1", 0, False
