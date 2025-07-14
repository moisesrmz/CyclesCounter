Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL7\"
archivos = Array( _
    "59Z113-000-F.txt", _
    "59Z113-000-L.txt", _
    "59Z113-000-L2.txt", _
    "59Z118-C00-L.txt", _
    "59Z153-000-F.txt", _
    "59Z153-000-K.txt", _
    "AMZ005-000-A.txt", _
    "AMZ032-C00-C.txt", _
    "AMZ040-C00-D.txt", _
    "AMZ040-C00-D2.txt", _
    "AMZ040-C00-E.txt", _
    "AMZW25-000-A.txt", _
    "AMZW25-000-B.txt" _
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

MsgBox resultado, 64, "CYCLES COUNTER EOL7"

Set objFSO = Nothing
Set objReadFile = Nothing
