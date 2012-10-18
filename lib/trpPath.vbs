' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

''' MODULES '''
Import "trpRunner"


''' MAIN '''
Function trppListFolder(directory)
    fout = ""
    For Each f In fso.GetFolder(directory).Files
        If Not fout = "" Then fout = fout + " "
        fout = fout & fso.GetBaseName(f.Name)
    Next
    trppListFolder = fout
End Function

Sub trppListPathFolders()
    trprPrintStatus vbCrLf & _
                    vbCrLf & _
                    "Apps in path: " & trppListFolder(envPathApps & "\PATH") & vbCrLf & vbCrLf & _
                    "Tools in path: " & trppListFolder(envPathTools), False
End Sub

Sub trppSetupPathFolders()
    ' Add to PATH
    trprWriteLine "set PATH=%trpTOOLS%;%trpAPPS%\PATH;%PATH%"
End Sub
