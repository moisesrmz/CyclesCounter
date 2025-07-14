Option Explicit

'Declare variables 
Dim objFSO, objReadFile, contents

'Set Objects 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\59z118-c00.txt", 1, False)

'Read file contents 
contents = objReadFile.ReadAll
contents = contents+1
objReadFile.close

Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\59z118-c00.txt", 2, False)
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
