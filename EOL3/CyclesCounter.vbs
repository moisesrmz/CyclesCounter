Option Explicit

 
Dim objFSO, objReadFile, content1,content2,content3,content4,content5,content6,content7,content8,content9,content10,content11,content12,content13,content14,content15,x


Set objFSO = CreateObject("Scripting.FileSystemObject") 

Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z114-000-C.txt", 1, False)

content1 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z114-000-A.txt", 1, False)
 
content2 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z153-C00-E.txt", 1, False)

content3 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z153-000-A.txt", 1, False)
 
content4 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z153-C00-K.txt", 1, False)

content5 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z153-C00-I.txt", 1, False)

content6 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z231-000-C.txt", 1, False)

content7 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\59Z232-000-A.txt", 1, False)

content8 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\903-847P-5IS.txt", 1, False)

content9 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\C903-847P-5IS-1.txt", 1, False)

content10 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\C903-847P-5IS-2.txt", 1, False)

content11 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZ040-C00-A.txt", 1, False)

content12 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZW17-000-A.txt", 1, False)

content13 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZW29-000-C.txt", 1, False)

content14 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZW25-000-A.txt", 1, False)

content15 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZW25-000-D.txt", 1, False)

content16 = objReadFile.ReadAll
Set objReadFile = objFSO.OpenTextFile("\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL3\AMZW25-000-B.txt", 1, False)
objReadFile.close

x=msgbox("59Z114-000-C:		"& content1&vbCrLf&"59Z114-000-A:		"& content2&vbCrLf&"59Z153-C00-E:		"& content3&vbCrLf&"59Z153-000-A:		"& content4&vbCrLf&"59Z153-C00-K:		"& content5&vbCrLf&"59Z231-000-C:		"& content6&vbCrLf&"59Z232-000-A:		"& content7&vbCrLf&"903-847P-5IS:		"& content8&vbCrLf&"C903-847P-5IS-1:		"& content9&vbCrLf&"C903-847P-5IS-2:		"& content10&vbCrLf&"AMZ040-C00-A:		"& content11&vbCrLf&"AMZW17-000-A:		"& content12&vbCrLf&"AMZW29-000-C:		"& content13&vbCrLf&"AMZW17-000-A:		"& content14&vbCrLf&"AMZW29-000-C:		"& content15&vbCrLf&"AMZW25-000-A:		"& content16&vbCrLf&"AMZW25-000-D:		"& content17&vbCrLf&"AMZW25-000-B:		"& content18&vbCrLf,64,"Cycles counter")

Set objFSO = Nothing 
Set objReadFile = Nothing

