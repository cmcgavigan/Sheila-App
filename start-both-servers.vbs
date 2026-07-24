' start-both-servers.vbs — launch BOTH receipt servers at once, hidden, each with
' its own tray icon:
'    Receipts  -> GREEN tray  (C:\Users\CMcGavigan\Documents\Receipts)
'    Sheila    -> GOLD  tray  (C:\Users\CMcGavigan\Documents\Sheilas app)
'
' Each app keeps its own folder, .env, port and Excel file — they're independent
' servers, just started together. This is what the Windows Startup shortcut runs.
'
' If you ever move either folder, update the two paths below.

Dim ws, fso
Set ws  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim receiptsLaunch, sheilaLaunch
receiptsLaunch = "C:\Users\CMcGavigan\Documents\Receipts\launch.vbs"
sheilaLaunch   = "C:\Users\CMcGavigan\Documents\Sheilas app\launch.vbs"

' Start the Receipts server (green) if its launcher exists.
If fso.FileExists(receiptsLaunch) Then
    ws.Run "wscript """ & receiptsLaunch & """", 0, False
End If

' Small stagger so the two Tailscale-serve calls don't race each other.
WScript.Sleep 2500

' Start Sheila's server (gold) if its launcher exists.
If fso.FileExists(sheilaLaunch) Then
    ws.Run "wscript """ & sheilaLaunch & """", 0, False
End If
