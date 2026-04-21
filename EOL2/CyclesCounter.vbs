Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase
Dim lineas, linea
Dim N1, M1, M2, maxValor
Dim partes

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL2\"

archivos = Array( _
    "59Z113-000-C.txt", "59Z113-000-C2.txt", "59Z113-000-D.txt", "59Z113-000-F.txt", "59Z113-000-L.txt", _
    "59Z117-C01-A.txt", "59Z118-C00-C.txt", "59Z118-C00-D.txt", "59Z118-C00-E.txt", "59Z118-C00-K.txt", "59Z118-C00-L.txt", _
    "59Z153-000-F.txt","59Z153-C00-A.txt", "59Z153-C00-A2.txt", _
    "59Z163-003-A.txt", "59Z163-003-B.txt", "59Z163-003-C.txt", "59Z163-003-D.txt", "59Z163-003-F.txt", _
    "59Z176-C01-A.txt", "59Z176-C01-B.txt", "59Z176-C01-C.txt", "59Z176-C01-D.txt", _
    "AMK12A-102Z5.txt", "AMZ005-000-F.txt", "AMZW01-000-A.txt", "AMZW01-000-C.txt", "AMZW17-000-C.txt" _
)

Set objFSO = CreateObject("Scripting.FileSystemObject")

resultado = ""

For i = 0 To UBound(archivos)

    nombreBase = objFSO.GetBaseName(archivos(i))
    contenido = ""

    If objFSO.FileExists(rutaBase & archivos(i)) Then

        Set objReadFile = objFSO.OpenTextFile(rutaBase & archivos(i), 1, False)
        contenido = Trim(objReadFile.ReadAll)
        objReadFile.Close

        ' Detectar formato nuevo o viejo
        If InStr(contenido, "=") > 0 Then
        
            ' Formato nuevo
            N1 = 0
            M1 = 0
            M2 = 0

            lineas = Split(contenido, vbCrLf)

            For Each linea In lineas

                If InStr(linea, "=") > 0 Then

                    partes = Split(linea, "=")

                    If UCase(Trim(partes(0))) = "N1" Then
                        N1 = CLng(partes(1))
                    ElseIf UCase(Trim(partes(0))) = "M1" Then
                        M1 = CLng(partes(1))
                    ElseIf UCase(Trim(partes(0))) = "M2" Then
                        M2 = CLng(partes(1))
                    End If

                End If

            Next

            ' Obtener valor mayor
            maxValor = N1
            If M1 > maxValor Then maxValor = M1
            If M2 > maxValor Then maxValor = M2

            contenido = maxValor

        Else

            ' Formato viejo
            contenido = CLng(contenido)

        End If

    Else
        contenido = "No se pudo leer"
    End If

    resultado = resultado & nombreBase & ":" & vbTab & contenido & vbCrLf

Next

resultado = UCase(resultado)

MsgBox resultado, 64, "CYCLES COUNTER EOL2"

Set objFSO = Nothing
Set objReadFile = Nothing