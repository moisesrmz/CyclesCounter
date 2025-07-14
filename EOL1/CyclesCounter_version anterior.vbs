Option Explicit

 
Dim objFSO, objReadFile, content1,content2,content3,content4,content5,content6,content7,content8,x


Set objFSO = CreateObject("Scripting.FileSystemObject") 

Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z118-c00.txt", 1, False)

content1 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z113-000.txt", 1, False)
 
content2 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z113-000a.txt", 1, False)

content3 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z114-000.txt", 1, False)
 
content4 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z163-003.txt", 1, False)

content5 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z163-003a.txt", 1, False)

content6 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\59z176-C01.txt", 1, False)

content7 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL1\AMK18A-102Z5.txt", 1, False)

content8 = objReadFile.ReadAll

objReadFile.close

x=msgbox("59z118-c00:	"& content1&vbCrLf&"59z113-000:	"& content2&vbCrLf&"59z113-000A:	"& content3&vbCrLf&"59z114-000:	"& content4&vbCrLf&"59z163-003:	"& content5&vbCrLf&"59z163-003A:	"& content6&vbCrLf&"59z176-C01:           "& content7&vbCrLf&"AMK18A-102Z5:           "& content8&vbCrLf,64,"Cycles counter")

Set objFSO = Nothing 
Set objReadFile = Nothing

