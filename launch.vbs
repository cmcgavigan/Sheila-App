' launch.vbs — starts the Sheila server with NO visible console window.
' Used by the Startup-folder shortcut so the server runs quietly in the
' background whenever Windows starts. Logs go to server.log in this folder.
Dim ws, fso, dir, nodeExe
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
ws.CurrentDirectory = dir

' Use a bundled node\node.exe if present, else node from PATH.
If fso.FileExists(dir & "\node\node.exe") Then
    nodeExe = dir & "\node\node.exe"
Else
    nodeExe = "node"
End If

' Run the PINK tray launcher (tray.js spawns server.js + shows the tray icon),
' redirecting output to server.log. Window style 0 = hidden.
ws.Run "cmd /c """"" & nodeExe & """ """ & dir & "\tray.js"" >> """ & dir & "\server.log"" 2>&1""", 0, False
