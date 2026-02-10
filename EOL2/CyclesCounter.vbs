Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

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

resultado = UCase(resultado) ' <-- Todo en mayúsculas

MsgBox resultado, 64, "CYCLES COUNTER EOL2"

Set objFSO = Nothing
Set objReadFile = Nothing
