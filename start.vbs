' tiaRoamingProfile v0.1.110519
' Created by Andras Tim @ 2011

Const setsPathApps = "apps"
Const setsPathData = "data"
Const setsLang = "HU"

''' COMMON RESOURCES '''
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim wshShell: Set wshShell = CreateObject( "WScript.Shell" )
Dim prgPath: prgPath = fso.GetParentFolderName(wscript.ScriptFullName)

Dim modules: Set modules = CreateObject("Scripting.Dictionary")
Dim moduleClassTemplate: moduleClassTemplate = _
"Class %CLASS%" & vbCrLf & _
"%CODE%" & vbCrLf & _
"Dim libPath" & vbCrLf & _
"Private Sub Class_Initialize()" & vbCrLf & _
"    libPath = ""%PATH%""" & vbCrLf & _
"    Dim existInit: On Error Resume Next: IsObject Module_Initialize: IsObject Module_Initialize: existInit = (Err.Number = 13): Err.Clear: On Error Goto 0" & vbCrLf & _
"    If existInit Then Module_Initialize" & vbCrLf & _
"End Sub" & vbCrLf & _
"End Class"

Public Sub Import(libName)
    Dim libFullPath: libFullPath =  prgPath & "\lib\" & libName & ".vbs"
    If Not modules.Exists(libName) Then
        'Load class contain from file
        Dim strCode, fo: Set fo = fso.OpenTextFile(libFullPath): strCode = fo.ReadAll: fo.Close
        'Get a new class template and substitue it
        Dim classCode: classCode = moduleClassTemplate
        classCode = Replace(classCode, "%CLASS%", "mod_" & libName)
        classCode = Replace(classCode, "%PATH%", fso.GetParentFolderName(libFullPath))
        classCode = Replace(classCode, "%CODE%", strCode)
        'Eval class
        On Error Resume Next
        ExecuteGlobal classCode
        If Not Err.Number = 0 Then
            WScript.StdOut.WriteLine "====== CODE ======" & vbCrLf & classCode & vbCrLf & "====== CODE ======" & vbCrLf
            WScript.StdOut.WriteLine "Microsoft VBScript runtime error: " & Err.Description & vbCrLf
            WScript.StdOut.WriteLine "MODULE: " & libName & vbCrLf & "METHOD: Code evaluate"
            WScript.Quit(1)
        End If
        On Error Goto 0
        'Register a new instance of class into modules
        ExecuteGlobal "modules.Add """ & libName & """, New mod_" & libName
        If Not Err.Number = 0 Then
            WScript.StdOut.WriteLine "MODULE: " & libName & vbCrLf & "METHOD: Module_Initialize"
            WScript.Quit(1)
        End If
    End If
End Sub

Public Sub Include(fileName)
    'Load VBScript from file
    Dim strCode, fo: Set fo = fso.OpenTextFile(prgPath & "\" & fileName): strCode = fo.ReadAll: fo.Close
    Execute strCode
End Sub


''' MODULES '''
Import "trpRunner"
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

    ' Overwrite parameters
    trpseSetEnv

    ' Set working directory
    trprWriteLine("%HOMEDRIVE%")
    trprWriteLine("cd """ & envPathData & """")
    trprPrintStatus " done", True

    ' Get and write program
    Dim runProgram: runProgram = Replace(Args2String, "`", "%")
    If runProgram = "" Then
        runProgram = "cmd"
        trppListPathFolder
    Else
        trprPrintStatus " * Starting program...", False
    End If
    trprPrintStatus "", False
    trprWriteLine(runProgram)

    ' Close file
    trprEndFile
End Sub


''' MAIN '''
Dim envPathApps: envPathApps = prgPath & "\" & setsPathApps
Dim envPathData: envPathData = prgPath & "\" & setsPathData
CreateCmd
trprRunAndClean

''' DEINIT '''
Set fso = Nothing
Set wshShell = Nothing
