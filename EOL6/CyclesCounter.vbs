Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL6\"
archivos = Array( _
    "59Z113-000-C.txt", _
    "59Z113-000-D.txt", _
    "59Z113-000-F.txt", _
    "59Z115-000-A.txt", _
    "59Z115-000-C.txt", _
    "59Z118-C00-L.txt", _
    "59Z120-C00-C.txt", _
    "59Z120-C00-C2.txt", _
    "59Z120-C00-D.txt", _
    "59Z153-000-C.txt", _
    "59Z153-000-D.txt", _
    "59Z153-000-L.txt", _
    "59Z153-C00-F.txt", _
    "59Z231-C00-A.txt", _
    "AMZ040-C00-A.txt", _
    "AMZ040-C00-B.txt", _
    "AMZW25-000-A.txt", _
    "AMZW25-000-A2.txt", _
    "AMZW25-000-B.txt", _
    "AMZW25-000-E.txt" _
)

Set objFSO = CreateObject("Scripting.FileSystemObject")

resultado = ""
For i = 0 To UBound(archivos)
    nombreBase = objFSO.GetBaseName(archivos(i))
    contenido = ""
    On Error Resume Next
    If objFSO.FileExists(rutaBase & archivos(i)) Then
        Set objReadFile = objFSO.OpenTextFile(rutaBase & archivos(i), 1, False)
        contenido = objReadFile.ReadAll
        objReadFile.Close
    Else
        contenido = "No se pudo leer"
    End If
    On Error GoTo 0
    resultado = resultado & nombreBase & ":" & vbTab & contenido & vbCrLf
Next

resultado = UCase(resultado)

MsgBox resultado, 64, "CYCLES COUNTER EOL6"

Set objFSO = Nothing
Set objReadFile = Nothing
