' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

''' MODULES '''
Import "trpPath"
Import "trpRunner"


''' MAIN '''
Sub trpseAddPath(envID, envPath)
    trprWriteLine("if not exist """ & envPath & """ mkdir """ & envPath & """" & vbCrLf & _
                  "set " & envID & "=" & envPath)
    trprPrintStatus ".", True
End Sub

Sub trpseSetEnv()
    Dim prgDrive: prgDrive = fso.GetDriveName(prgPath)
    trprWriteLine "set HOMEDRIVE=" & prgDrive
    trprWriteLine "set HOMEPATH=" & Right(envPathData, Len(envPathData) - Len(prgDrive)) & "\USER"

    trpseAddPath "trpAPPS", envPathApps
    trpseAddPath "trpDATA", envPathData

    trpseAddPath "USERPROFILE",     envPathData & "\USER"
    trpseAddPath "LOCALAPPDATA",     envPathData & "\USER\AppData\Local"
    trpseAddPath "APPDATA",         envPathData & "\USER\AppData\Roaming"
    trpseAddPath "PUBLIC",             envPathData & "\PUBLIC"
    trpseAddPath "ALLUSERSPROFILE",    envPathData & "\PROGRAMDATA"
    trpseAddPath "ProgramData",     envPathData & "\PROGRAMDATA"

    trprWriteLine "set LANG=" & setsLang

    trppSetupPathFolder
End Sub
