'tiaRoamingProfile v0.1.5
'Created by Andras Tim @ 2011

Const envLang = "HU"
Const envPathData = "data"
Const envPathApps = "apps"
Const envPathLib = "lib"

dim prgPath, prevStatText
prevStatText = ""

Set wshShell = CreateObject( "WScript.Shell" )
Set fso = CreateObject("Scripting.FileSystemObject")


''' SUBS, FUNCTIONS '''
Sub PrintStatus(fd, text, continuePrevLine)
    ' Relpace new lines
    curr = Replace(text, vbCrLf, vbCrLf & "echo ")

    ' Like "echo -n"
    If continuePrevLine Then
        curr = curr
    Else
        curr = vbCrLf & "echo " & curr
    End If

    ' Replace empyt lines
    aCurr = Split(curr, vbCrLf)
    For i = 0 to UBound(aCurr)
        If Trim(aCurr(i)) = "echo" Then aCurr(i) = "echo."
    Next
    curr = Join(aCurr, vbCrLf)

    ' Write out
    fd.WriteLine("cls" & vbCrLf & _
                 "echo ..:: tiaRoamingProfile ::.. " & vbCrLf & _
                 "echo." & prevStatText & curr)
    prevStatText = prevStatText & curr
End Sub

Sub ChangeEnvFolder(fd, envID, envPath)
    fd.WriteLine("if not exist """ & envPath & """ mkdir """ & envPath & """" & vbCrLf & _
                 "set " & envID & "=" & envPath)
    PrintStatus of, ".", True
End Sub

Function GetArgPrg()
    ret = ""
    If (WScript.Arguments.Count > 0) Then
        For i = 0 to WScript.Arguments.Count - 1
            If Not i = 0 Then ret = ret & " "
            ret = ret & """" & WScript.Arguments.Item(i) & """"
        Next
        ret = Replace(ret, "`", "%")
    End If
    GetArgPrg = ret
End Function

Sub DirPATH(fd)
    fout = ""
    For Each f In fso.GetFolder(prgPath + "\" & envPathApps & "\PATH").Files
        If Not fout = "" Then fout = fout + " "
        fout = fout & fso.GetBaseName(f.Name)
    Next
    PrintStatus fd, vbCrLf & vbCrLf &  "Apps in path: " & fout, False
End Sub


''' INIT '''
' Get paths
prgPath = fso.GetParentFolderName(wscript.ScriptFullName)
envHomeDrive = fso.GetDriveName(prgPath)
envHomePath = Right(prgPath, Len(prgPath) - Len(prgDrive)) & "\" & envPathData
tmpPath = wshShell.ExpandEnvironmentStrings( "%TEMP%" )

' Get prg temporary file name
Randomize
cmdPath = tmpPath & "\tiaRoamingProfile_" & Fix(Rnd*100000) & ".cmd"


''' PREPARE STARTER CMD '''
' Open file
Set of = fso.CreateTextFile(cmdPath, True)
of.WriteLine("@echo off")
PrintStatus of, " * Preparing enviroment", False

' Overwrite parameters
ChangeEnvFolder of, "trpAPPS", prgPath + "\" & envPathApps
ChangeEnvFolder of, "trpLIB", prgPath + "\" & envPathLib
ChangeEnvFolder of, "trpREPO", prgPath + "\" & envPathData & "\REPO"
ChangeEnvFolder of, "USERPROFILE", prgPath & "\" & envPathData
ChangeEnvFolder of, "LOCALAPPDATA", prgPath & "\" & envPathData & "\APPDATA\Local"
ChangeEnvFolder of, "APPDATA", prgPath & "\" & envPathData & "\APPDATA\Roaming"
ChangeEnvFolder of, "PUBLIC", prgPath & "\" & envPathData & "\PUBLIC"
ChangeEnvFolder of, "ALLUSERSPROFILE", prgPath & "\" & envPathData & "\PROGRAMDATA"
ChangeEnvFolder of, "ProgramData", prgPath & "\" & envPathData & "\PROGRAMDATA"
of.WriteLine("set HOMEDRIVE=" & envHomeDrive &  vbCrLf & _
             "set HOMEPATH=" & envHomePath  & vbCrLf & _
             "set LANG=" & envLang & vbCrLf & _
             "set PATH=%trpAPPS%\PATH;%PATH%")

' Set working directory
of.WriteLine(prgDrive)
of.WriteLine("cd ""%USERPROFILE%""")
PrintStatus of, " done", True

' Get run program and start
runProgram = GetArgPrg
If runProgram = "" Then
    runProgram = "cmd"
    DirPATH of
Else
    PrintStatus of, " * Starting program...", False
End If
PrintStatus of, "", False
of.WriteLine(runProgram)

' Exit
of.WriteLine("exit")
of.Close

''' START SCRIPT '''
wshShell.Run cmdPath, 1, True

''' EXIT '''
' Remove temporary files
Set of = fso.GetFile(cmdPath)
of.Delete

Set of = Nothing
Set fso = Nothing
Set wshShell = Nothing
