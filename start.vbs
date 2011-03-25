'tiaRoamingProfile v0.1
'Created by Andras Tim @ 2011

dim prgPath

Set wshShell = CreateObject( "WScript.Shell" )
Set fso = CreateObject("Scripting.FileSystemObject")

''' INIT '''
' Get paths
prgPath = fso.GetParentFolderName(wscript.ScriptFullName)
prgDrive = fso.GetDriveName(prgPath)
tmpPath = wshShell.ExpandEnvironmentStrings( "%TEMP%" )

' Get prg temporary file name
Randomize
cmdPath = tmpPath & "\tiaRoamingDevel_" & Fix(Rnd*100000) & ".cmd"

' Get run program
If (WScript.Arguments.Count > 0) Then
    runProgram = fso.GetAbsolutePathName(wshShell.ExpandEnvironmentStrings(WScript.Arguments.Item(0)))
Else
    runProgram = "cmd"
End If


''' PREPARE STARTER CMD '''
' Open file
Set of = fso.CreateTextFile(cmdPath, True)
of.WriteLine("@echo off")
' Overwrite parameters
of.WriteLine("set HOMEDRIVE=" & prgDrive)
of.WriteLine("set HOMEPATH=" & prgPath)
of.WriteLine("set USERPROFILE=" & prgPath & "\profile")
of.WriteLine("set ALLUSERSPROFILE=" & prgPath & "\profile")
of.WriteLine("set APPDATA=" & prgPath & "\profile\APPDATA")
' Set working directory
of.WriteLine(prgDrive)
of.WriteLine("cd """ & prgPath & "\repo""")
' Start program
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