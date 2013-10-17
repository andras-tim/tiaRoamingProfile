' tiaRoamingProfile v0.1.131017
' Created by Andras Tim @ 2013

''' MODULES '''
Import "trpSessionManager"
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
    Dim sessionDir: sessionDir = trpsmGetSessionDirectory
    trprWriteLine "set HOMEDRIVE=" & prgDrive
    trprWriteLine "set HOMEPATH=" & Right(envPathData, Len(envPathData) - Len(prgDrive)) & "\USER"

    trpseAddPath "trpAPPS", envPathApps
    trpseAddPath "trpDATA", envPathData
    trpseAddPath "trpTOOLS", envPathTools
    trprWriteLine "set trpSESSION=" & sessionDir

    trpseAddPath "USERPROFILE",     envPathData & "\USER"
    trpseAddPath "LOCALAPPDATA",     envPathData & "\USER\AppData\Local"
    trpseAddPath "APPDATA",         envPathData & "\USER\AppData\Roaming"
    trpseAddPath "PUBLIC",             envPathData & "\PUBLIC"
    trpseAddPath "ALLUSERSPROFILE",    envPathData & "\PROGRAMDATA"
    trpseAddPath "ProgramData",     envPathData & "\PROGRAMDATA"

    trprWriteLine "set LANG=" & setsLang

    trppSetupPathFolders
End Sub
