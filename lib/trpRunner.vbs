' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

''' MODULES '''
Import "trpSessionManager"


''' MAIN '''
Dim trprPrevStatText: trprPrevStatText = ""
Dim trprFo, trprCmdPath

Sub trprInitFile()
    trprCmdPath = trpsmGetSessionDirectory & "\session.cmd"
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

Function trprRun()
    trprRun = wshShell.Run(trprCmdPath, 1, True)
End Function
