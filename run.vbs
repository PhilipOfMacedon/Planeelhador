Set shell = CreateObject("WScript.Shell")
' The %1 in the VBScript is the parameter passed from the Registry (%V)
' The "0" ensures the command runs hidden (without a window)
shell.Run "cmd /c cd /d " & WScript.Arguments(0) & " && python C:\Planeelhador\Planeelhador_support.py " & WScript.Arguments(0), 0, False
Set shell = Nothing