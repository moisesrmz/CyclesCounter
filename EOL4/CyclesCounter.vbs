Option Explicit

Dim objFSO, objReadFile
Dim archivos, i, resultado, contenido
Dim rutaBase, nombreBase

rutaBase = "\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\EOL4\"
archivos = Array( _
    "1-2291859-2.txt", _
    "2291859-1.txt", _
    "59Z113-000-F.txt", _
    "59Z113-000-N.txt", _
    "59Z114-000-A.txt", _
    "59Z114-000-C.txt", _
    "59Z118-C00-A.txt", _
    "59Z118-C00-F.txt", _
    "59Z153-C00-A.txt", _
    "59Z163-003-E.txt", _
    "59Z176-C01-F.txt", _
    "59Z232-000-C.txt", _
    "903-847P-51S.txt", _
    "AMK12A-102Z5.txt", _
    "AMZ010-C00-A.txt", _
    "AMZ010-C00-B.txt", _
    "AMZ010-C00-C.txt", _
    "AMZ010-C00-F.txt", _
    "AMZ025-000-D.txt", _
    "AMZ032-C00-B.txt", _
    "AMZW01-000-F.txt", _
    "AMZW17-000-C.txt", _
    "AMZW31-000-B.txt" _
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

MsgBox resultado, 64, "CYCLES COUNTER EOL4"

Set objFSO = Nothing
Set objReadFile = Nothing
