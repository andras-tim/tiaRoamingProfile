' tiaRoamingProfile v0.1.131017
' Created by Andras Tim @ 2013

Const setsPathApps = "apps"
Const setsPathData = "data"
Const setsPathTools = "tools"
Const setsPathVendors = "vendors"
Const setsLang = "HU"

''' COMMON RESOURCES '''
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim wshShell: Set wshShell = CreateObject( "WScript.Shell" )
Dim prgPath: prgPath = fso.GetParentFolderName(wscript.ScriptFullName)

Dim modules: Set modules = CreateObject("Scripting.Dictionary")
Sub Import(modulePath)
    Dim libFullPath: libFullPath =  prgPath & "\lib\" & modulePath & ".vbs"
    Dim libBaseName: libBaseName =  fso.GetBaseName(libFullPath)
    If Not modules.Exists(libBaseName) Then
        Dim strCode, fo: Set fo = fso.OpenTextFile(libFullPath): strCode = fo.ReadAll: fo.Close
        modules.Add libBaseName, Array(modulePath & ".vbs", fso.GetParentFolderName(libFullPath))
        ExecuteGlobal strCode
    End If
End Sub


''' MODULES '''
Import "trpSessionManager"
Import "trpRunner"
Import "trpPreConf"
Import "trpPath"
Import "trpSetEnv"


''' SUBS, FUNCTIONS '''
Function Args2String()
    ret = ""
    If (WScript.Arguments.Count > 0) Then
        For i = 0 to WScript.Arguments.Count - 1
            If Not i = 0 Then ret = ret & " "
            ret = ret & """" & WScript.Arguments.Item(i) & """"
        Next
    End If
    Args2String = ret
End Function

Sub CreateCmd()
    ' Prepare starter cmd
    trprInitFile
    trprPrintStatus " * Preparing enviroment", False

    ' Setup environment varaibles, directories, symlinks
    trpseSetupEnvironment

    ' Set working directory
    trprWriteLine("%HOMEDRIVE%")
    trprWriteLine("cd """ & envPathData & """")
    trprPrintStatus " done", True

    ' Get and write program
    Dim runProgram: runProgram = Replace(Args2String, "`", "%")
    If runProgram = "" Then
        runProgram = "cmd"
        trppListPathFolders
        trprPrintStatus "", False
    Else
        trprPrintStatus " * Starting program...", False
    End If
    trprPrintStatus "", False
    trprWriteLine(runProgram)

    ' Close file
    trprEndFile
End Sub


''' MAIN '''
Dim envPathTools: envPathTools = prgPath & "\" & setsPathTools
Dim envPathApps: envPathApps = prgPath & "\" & setsPathApps
Dim envPathData: envPathData = prgPath & "\" & setsPathData
Dim envPathVendors: envPathVendors = prgPath & "\" & setsPathVendors
Dim mainRetCode: mainRetCode = 0

trpsmInitializeSessionDirectory
trppcPrepare
CreateCmd
mainRetCode = trprRun
trpsmPurgeSessionDirectory

''' DEINIT '''
Set fso = Nothing
Set wshShell = Nothing

WScript.Quit mainRetCode
