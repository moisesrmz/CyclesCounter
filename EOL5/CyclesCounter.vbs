Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL5\"
archivos = Array( _
    "59Z113-000-C.txt", _
    "59Z113-000-F.txt", _
    "59Z114-000-A.txt", _
    "59Z117-C01-A.txt", _
    "59Z118-C00-D.txt", _
    "59z118-C00-F.txt", _
    "59Z153-000-A.txt", _
    "59Z153-000-B.txt", _
    "59Z153-000-D.txt", _
    "59Z153-C00-B.txt", _
    "59Z153-C00-F.txt", _
    "59Z153-C00-I.txt", _
    "59Z163-003-A.txt", _
    "59Z163-003-B.txt", _
    "59Z163-003-B2.txt", _
    "59Z163-003-F.txt", _
    "59Z176-C01-B.txt", _
    "59Z176-C01-B2.txt", _
    "59Z176-C01-F.txt", _
    "59Z178-000-F.txt", _
    "59Z231-000-A.txt", _
    "59Z232-000-A.txt", _
    "AMZ040-C00-A.txt", _
    "AMZ040-C00-B.txt", _
    "AMZ14P-001-A.txt", _
    "AMZW17-000-A.txt", _
    "AMZW17-000-B.txt", _
    "AMZW17-000-N.txt", _
    "AMZW25-000-B.txt", _
    "AMZW29-000-C.txt" _
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

MsgBox resultado, 64, "CYCLESCOUNTER EOL5"

Set objFSO = Nothing
Set objReadFile = Nothing
