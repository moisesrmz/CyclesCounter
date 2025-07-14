Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\Percha2\"
archivos = Array( _
    "59Z113-000-C.txt", _
    "59Z113-000-F.txt", _
    "59Z113-000-L.txt", _
    "59Z117-C01-A.txt", _
    "59Z118-C00-D.txt", _
    "59Z118-C00-L.txt", _
    "59Z163-003-A.txt", _
    "59Z163-003-B.txt", _
    "59Z176-C01-B.txt", _
    "59Z176-C01-C.txt", _
    "59Z176-C01-D.txt" _
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

MsgBox resultado, 64, "CYCLES COUNTER PERCHA2"

Set objFSO = Nothing
Set objReadFile = Nothing
