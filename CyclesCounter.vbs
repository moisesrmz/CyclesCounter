Option Explicit

Dim objFSO, objReadFile
Dim rutaBase, testerName
Dim resultado
Dim carpeta, archivo
Dim contenido
Dim lineas, linea
Dim N1, M1, M2, maxValor
Dim partes

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Leer tester actual
Dim testerFile
testerFile = "C:\Users\Public\Documents\Cirris\tester.txt"

If Not objFSO.FileExists(testerFile) Then
    MsgBox "tester.txt no encontrado", 16
    WScript.Quit
End If

Set objReadFile = objFSO.OpenTextFile(testerFile,1)
testerName = Trim(objReadFile.ReadAll)
objReadFile.Close

' Construir ruta automaticamente
rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\" & testerName & "\"

If Not objFSO.FolderExists(rutaBase) Then
    MsgBox "No se encontro la carpeta del tester: " & rutaBase,16
    WScript.Quit
End If

Set carpeta = objFSO.GetFolder(rutaBase)

resultado = ""

' Recorrer todos los TXT (modulos)
For Each archivo In carpeta.Files

    If LCase(objFSO.GetExtensionName(archivo.Name)) = "txt" Then

        contenido = ""

        Set objReadFile = objFSO.OpenTextFile(archivo.Path,1)
        contenido = Trim(objReadFile.ReadAll)
        objReadFile.Close

        ' Detectar formato nuevo o viejo
        If InStr(contenido,"=") > 0 Then

            N1 = 0
            M1 = 0
            M2 = 0

            lineas = Split(contenido,vbCrLf)

            For Each linea In lineas

                If InStr(linea,"=") > 0 Then

                    partes = Split(linea,"=")

                    If UCase(Trim(partes(0))) = "N1" Then
                        N1 = CLng(partes(1))
                    ElseIf UCase(Trim(partes(0))) = "M1" Then
                        M1 = CLng(partes(1))
                    ElseIf UCase(Trim(partes(0))) = "M2" Then
                        M2 = CLng(partes(1))
                    End If

                End If

            Next

            maxValor = N1
            If M1 > maxValor Then maxValor = M1
            If M2 > maxValor Then maxValor = M2

            contenido = maxValor

        Else

            contenido = CLng(contenido)

        End If

        resultado = resultado & UCase(objFSO.GetBaseName(archivo.Name)) & ":" & vbTab & contenido & vbCrLf

    End If

Next

MsgBox resultado,64,"CYCLES COUNTER - " & testerName

Set objFSO = Nothing
Set objReadFile = Nothing