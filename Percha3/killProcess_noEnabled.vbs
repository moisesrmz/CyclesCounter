'Option Explicit
'strCommand = "taskkill /f /im CalculatorApp.exe"
'x=msgbox("Se han rebasado las 100,000 activaciones de la contraparte 'XXX', favor de contactar a Ingenieria de Pruebas.",16,"Limite Rebasado")

'Set WshShell = CreateObject("WScript.Shell")
'Set WshShellExec = WshShell.Exec(strCommand)
 

contents = 1000
If contents > 100000 Then
	strCommand = "taskkill /f /im CalculatorApp.exe"
	x=msgbox("Se han rebasado las 100,000 activaciones de la contraparte '59Z118-C00', favor de contactar a Ingenieria de Pruebas.",16,"Limite Rebasado")
	Set WshShell = CreateObject("WScript.Shell")
	Set WshShellExec = WshShell.Exec(strCommand)
Else
	x=msgbox("Toro bien.",32,"Toro bien")

End If