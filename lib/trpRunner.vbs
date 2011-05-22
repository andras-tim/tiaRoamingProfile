' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

''' MODULES '''


''' MAIN '''
Dim trprPrevStatText: trprPrevStatText = ""
Dim trprFo, trprCmdPath

Sub trprInitFile()
    Dim tmpPath: tmpPath = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    Randomize
    trprCmdPath = tmpPath & "\tiaRoamingProfile_" & Fix(Rnd*100000) & ".cmd"
    Set trprFo = fso.CreateTextFile(trprCmdPath, True)
    trprWriteLine("@echo off")
End Sub

Sub trprEndFile()
    trprWriteLine("exit")
    trprFo.Close
End Sub

Sub trprWriteLine(text)
    trprFo.WriteLine(text)
End Sub


Sub trprPrintStatus(text, continuePrevLine)
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
    trprWriteLine("cls" & vbCrLf & _
                  "echo ..:: tiaRoamingProfile ::.. " & vbCrLf & _
                  "echo." & trprPrevStatText & curr)
    trprPrevStatText = trprPrevStatText & curr
End Sub

Sub trprRunAndClean()
    'Start cmd
    wshShell.Run trprCmdPath, 1, True
    'Delete temp file
    Dim fo: Set fo = fso.GetFile(trprCmdPath)
    fo.Delete
End Sub
