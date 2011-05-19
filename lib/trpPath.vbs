' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

Sub trppListPathFolder()
    fout = ""
    For Each f In fso.GetFolder(envPathApps & "\PATH").Files
        If Not fout = "" Then fout = fout + " "
        fout = fout & fso.GetBaseName(f.Name)
    Next
    trprPrintStatus vbCrLf & vbCrLf &  "Apps in path: " & fout, False
End Sub

Sub trppSetupPathFolder()
    ' Add to PATH
    trprWriteLine "set PATH=%trpAPPS%\PATH;%PATH%"
End Sub
