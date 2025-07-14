'Set WshShell = WScript.CreateObject("WScript.Shell")
'Return = WshShell.Run("runCamera.bat", 1, true)

'strCommand = "cmd /K start microsoft.windows.camera:"
'Set WshShell = CreateObject("WScript.Shell")
'Set WshShellExec = WshShell.Exec(strCommand)




'Set WshShellExec = WshShell.Exec("runCamera.bat")
'Dim oShell
'Set oShell = WScript.CreateObject ("WScript.Shell")
'oShell.run "cmd /K start microsoft.windows.camera:"
'Set oShell = Nothing

' Run one command

'Call WshShell.Run("cmd /K start microsoft.windows.camera:")

'Const WshFinished = 1
'Const WshFailed = 2
strCommand = "cmd /K start microsoft.windows.camera:&&exit"

Set WshShell = CreateObject("WScript.Shell")
Set WshShellExec = WshShell.Exec(strCommand)

'Select Case WshShellExec.Status
'   Case WshFinished
'       strOutput = WshShellExec.StdOut.ReadAll
'   Case WshFailed
'       strOutput = WshShellExec.StdErr.ReadAll
'End Select

'WScript.StdOut.Write strOutput  'write results to the command line
'WScript.Echo "VERIFICAR CALIBRACION VIGENTE"          'write results to default output
'MsgBox strOutput                'write results in a message box