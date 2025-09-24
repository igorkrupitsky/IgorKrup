Option Explicit

Dim fso, folder, asmFile, file, lines, line, versionLine, i, versionParts, newVersion
Set fso = CreateObject("Scripting.FileSystemObject")

' Get folder of this script
Set folder = fso.GetFile(WScript.ScriptFullName).ParentFolder
asmFile = fso.BuildPath(folder.Path, "AssemblyInfo.vb")

If Not fso.FileExists(asmFile) Then
    MsgBox "AssemblyInfo.vb not found in: " & folder.Path, vbCritical, "Error"
    WScript.Quit 1
End If

' Read all lines
Set file = fso.OpenTextFile(asmFile, 1)
lines = Split(file.ReadAll, vbCrLf)
file.Close

' Look for AssemblyFileVersion
For i = 0 To UBound(lines)
    If InStr(lines(i), "AssemblyFileVersion") > 0 Then
        versionLine = lines(i)

        ' Extract version numbers
        Dim regex, matches
        Set regex = New RegExp
        regex.Pattern = "\d+\.\d+\.\d+\.\d+"
        regex.Global = False

        If regex.Test(versionLine) Then
            Set matches = regex.Execute(versionLine)
            versionParts = Split(matches(0).Value, ".")

            ' Increment last part (revision)
            versionParts(3) = CInt(versionParts(3)) + 1

            ' Build new version string
            newVersion = versionParts(0) & "." & versionParts(1) & "." & versionParts(2) & "." & versionParts(3)

            ' Replace line
            lines(i) = Replace(versionLine, matches(0).Value, newVersion)

            Exit For
        End If
    End If
Next

' Write file back
Set file = fso.OpenTextFile(asmFile, 2, True)
file.Write Join(lines, vbCrLf)
file.Close

MsgBox "AssemblyFileVersion incremented to " & newVersion, vbInformation, "Success"
