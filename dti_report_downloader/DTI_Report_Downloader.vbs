' DTI Portal Report Downloader

Dim scriptDir, pythonScript, shell
Dim fsoTemp
Set fsoTemp = CreateObject("Scripting.FileSystemObject")
' Resolve shortcut -> get real VBS path
Dim realPath
realPath = WScript.ScriptFullName
If fsoTemp.GetExtensionName(realPath) = "lnk" Or InStr(realPath, "Shortcut") > 0 Then
    Dim sh2
    Set sh2 = CreateObject("WScript.Shell")
    realPath = sh2.CreateShortcut(realPath).TargetPath
End If
scriptDir    = fsoTemp.GetParentFolderName(realPath)
pythonScript = scriptDir & "\dti_report.py"
Set shell    = CreateObject("WScript.Shell")

' Pokusaj sa "python" komandom
Dim result
result = shell.Run("cmd /k python """ & pythonScript & """", 1, True)

If result <> 0 Then
    ' Ako python nije u PATH, pokusaj sa punom putanjom
    Dim fso, pyPaths, p, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    pyPaths = Array( _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Python\pythoncore-3.14-64\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Python\pythoncore-3.14-64\Scripts\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Python\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python314\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\python.exe", _
        shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\python.exe", _
        "C:\Python314\python.exe", _
        "C:\Python313\python.exe" _
    )
    For i = 0 To UBound(pyPaths)
        If fso.FileExists(pyPaths(i)) Then
            shell.Run "cmd /k """ & pyPaths(i) & """ """ & pythonScript & """", 1, True
            WScript.Quit
        End If
    Next
    MsgBox "Python nije pronadjen!" & vbCrLf & vbCrLf & _
           "Pokreni setup.bat iz foldera:" & vbCrLf & scriptDir, 16, "DTI Report - Greska"
End If
