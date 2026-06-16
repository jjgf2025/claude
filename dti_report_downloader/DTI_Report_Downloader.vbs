' DTI Portal Report Downloader
' Pokretanje bez crnog konzolnog prozora

Dim scriptDir, pythonScript, shell
scriptDir   = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
pythonScript = scriptDir & "\dti_report.py"
Set shell = CreateObject("WScript.Shell")

' Pokreni Python skriptu u novom konzolnom prozoru (vidljiv progress)
shell.Run "cmd /c python """ & pythonScript & """", 1, False
