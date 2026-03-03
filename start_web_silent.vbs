
' 背景執行 web_app.py，不顯示視窗
Dim objShell
Set objShell = CreateObject("WScript.Shell")

Dim scriptDir
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

objShell.Run "pythonw """ & scriptDir & "web_app.py""", 0, False
