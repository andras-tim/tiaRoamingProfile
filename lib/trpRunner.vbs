' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

''' MODULES '''
Import "trpSessionManager"


''' MAIN '''
Dim trprPrevStatText: trprPrevStatText = ""
Dim trprFo, trprCmdPath
Dim trprFunctions: Set trprFunctions = CreateObject("Scripting.Dictionary")

trprAddFunction "trprPrintCmdHeader", _
                "cls" & vbCrLf & _
                "echo ..:: tiaRoamingProfile ::.. " & vbCrLf & _
                "echo."


''' SUBS, FUNCTIONS '''
Sub trprInitFile()
    trprCmdPath = trpsmGetSessionDirectory & "\session.cmd"
    Set trprFo = fso.CreateTextFile(trprCmdPath, True)
    trprWriteLine("@echo off")
End Sub

Sub trprEndFile()
    trprWriteLine("goto end")
    trprWriteFunctions
    trprWriteLine(":end" & vbCrLf & _
                  "exit" & vbCrLf)
    trprFo.Close
End Sub

Sub trprAddFunction(name, commands)
    trprFunctions.Add name, ":" & name & vbCrLf & _
                            commands & vbCrLf & _
                            "goto :eof"
End Sub

Sub trprCallFunction(name, txtConcatenatedParameters)
    trprWriteLine "call :" & name & " " & txtConcatenatedParameters
End Sub

Sub trprWriteLine(text)
    trprFo.WriteLine(text)
End Sub

Sub trprWriteFunctions()
    For Each name In trprFunctions
        trprWriteLine(trprFunctions(name))
    Next
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
    trprCallFunction "trprPrintCmdHeader", trprPrevStatText & curr
    trprPrevStatText = trprPrevStatText & curr
End Sub

Function trprRun()
    trprRun = wshShell.Run(trprCmdPath, 1, True)
End Function
