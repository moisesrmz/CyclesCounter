Option Explicit

 
Dim objFSO, objReadFile, content1,x


Set objFSO = CreateObject("Scripting.FileSystemObject") 

Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\Percha1\59Z113-000-F.txt", 1, False)
content1 = objReadFile.ReadAll



objReadFile.close

x=msgbox("59Z113-000-F:	"& content1&vbCrLf,64,"Cycles counter Percha1")

Set objFSO = Nothing 
Set objReadFile = Nothing

