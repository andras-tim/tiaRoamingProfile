' tiaRoamingProfile v0.1.131017
' Created by Andras Tim @ 2013

''' MODULES '''
Import "trpSessionManager"
Import "trpPath"
Import "trpRunner"


''' MAIN '''
' Reade more SID on http://support.microsoft.com/kb/243330
Const SID_EVERYONE = "S-1-1-0"
Const SID_LOCAL_SYSTEM = "S-1-5-18"

trprAddFunction "trpseCreateShortcutLikeLink", _
                "set link_path=%~1" & vbCrLf & _
                "set target_path=%~2" & vbCrLf & _
                "if not exist ""%target_path%"" mkdir ""%target_path%""" & vbCrLf & _
                "if exist ""%link_path%"" goto :eof" & vbCrLf & _
                "call ""%trpTOOLS%\symlinkDir.cmd"" ""%link_path%"" ""%target_path%""" & vbCrLf & _
                "icacls ""%link_path%"" /deny *" & SID_EVERYONE & ":(RD) /L > ""%DEVNULL_OUT%""" & vbCrLf & _
                "icacls ""%link_path%"" /setowner *" & SID_LOCAL_SYSTEM & " /L > ""%DEVNULL_OUT%""" & vbCrLf & _
                "attrib +S +H +I ""%link_path%"" /L"


''' SUBS, FUNCTIONS '''
Sub trpseCreateAndAddPath(envID, envPath)
    trprWriteLine("if not exist """ & envPath & """ mkdir """ & envPath & """" & vbCrLf & _
                  "set " & envID & "=" & envPath)
    trprPrintStatus ".", True
End Sub

Sub trpseCreateShortcutLikeLink(linkPath, targetPath)
    trprCallFunction "trpseCreateShortcutLikeLink", """" & linkPath & """ """ & targetPath & """"
    trprPrintStatus ".", True
End Sub

Sub trpseSetupEnvironment()
    Dim prgDrive: prgDrive = fso.GetDriveName(prgPath)
    Dim sessionDir: sessionDir = trpsmGetSessionDirectory

    trprWriteLine "set HOMEDRIVE=" & prgDrive
    trprWriteLine "set HOMEPATH=" & Right(envPathData, Len(envPathData) - Len(prgDrive)) & "\USER"

    trpseCreateAndAddPath "trpAPPS", envPathApps
    trpseCreateAndAddPath "trpDATA", envPathData
    trpseCreateAndAddPath "trpTOOLS", envPathTools
    trpseCreateAndAddPath "trpVENDORS", envPathVendors
    trprWriteLine "set trpSESSION=" & sessionDir

    trprWriteLine "set DEVNULL_OUT=" & sessionDir & "\devnull_out.txt"
    trprWriteLine "set DEVNULL_ERR=" & sessionDir & "\devnull_err.txt"

    trpseCreateAndAddPath "USERPROFILE", envPathData & "\USER"
    trpseCreateAndAddPath "LOCALAPPDATA", envPathData & "\USER\AppData\Local"
    trpseCreateAndAddPath "APPDATA", envPathData & "\USER\AppData\Roaming"
    trpseCreateAndAddPath "PUBLIC", envPathData & "\PUBLIC"
    trpseCreateAndAddPath "ALLUSERSPROFILE", envPathData & "\PROGRAMDATA"
    trpseCreateAndAddPath "ProgramData", envPathData & "\PROGRAMDATA"

    trpseCreateShortcutLikeLink envPathData & "\USER\Application Data", envPathData & "\USER\AppData\Roaming"
    trpseCreateShortcutLikeLink envPathData & "\USER\Cookies", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\Cookies"
    trpseCreateShortcutLikeLink envPathData & "\USER\Local Settings", envPathData & "\USER\AppData\Local"
    trpseCreateShortcutLikeLink envPathData & "\USER\AppData\Local\Application Data", envPathData & "\USER\AppData\Local"
    trpseCreateShortcutLikeLink envPathData & "\USER\AppData\Local\Temporary Internet Files", envPathData & "\USER\AppData\Local\Microsoft\Windows\Temporary Internet Files"
    trpseCreateShortcutLikeLink envPathData & "\USER\NetHood", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\Network Shortcuts"
    trpseCreateShortcutLikeLink envPathData & "\USER\PrintHood", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\Printer Shortcuts\"
    trpseCreateShortcutLikeLink envPathData & "\USER\SendTo", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\SendTo"
    trpseCreateShortcutLikeLink envPathData & "\USER\Start Menu", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\Start Menu"
    trpseCreateShortcutLikeLink envPathData & "\USER\Templates", envPathData & "\USER\AppData\Roaming\Microsoft\Windows\Templates"

    trprWriteLine "set LANG=" & setsLang

    trppSetupPathFolders
End Sub
