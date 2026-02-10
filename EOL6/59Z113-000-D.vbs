Option Explicit

'Declare variables 
Dim objFSO, objReadFile, contents, strCommand, x, WshShell, WshShellExec

'Set Objects 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL6\59Z113-000-D.txt", 1, False)

'Read file contents 
contents = objReadFile.ReadAll
contents = contents+1
objReadFile.close

'kill process start
If contents >= 30000 Then
	strCommand = "taskkill /f /im easywire.exe"
	x=msgbox("Se han rebasado las 30,000 activaciones de la contraparte '59Z113-000-D', favor de contactar a Ingenieria de Pruebas.",16,"Limite Rebasado")
	Set WshShell = CreateObject("WScript.Shell")
	Set WshShellExec = WshShell.Exec(strCommand)	
End If
'kill proces end

Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL6\59Z113-000-D.txt", 2, False)
objReadFile.Write  (contents)
'Close file 
objReadFile.close

'Display results 
'wscript.echo contents

'Cleanup objects 
Set objFSO = Nothing 
Set objReadFile = Nothing

'Quit script 
'WScript.Quit()
