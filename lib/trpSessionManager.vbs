' tiaRoamingProfile v0.1.20131017
' Created by Andras Tim @ 2013

''' MODULES '''


''' MAIN '''
Dim trpsmSessionDirectory: trpsmSessionDirectory = ""

Sub trpsmInitializeSessionDirectory()
	Dim tmpPath: tmpPath = wshShell.ExpandEnvironmentStrings("%TEMP%")
	Do
		Randomize
		trpsmSessionDirectory = tmpPath & "\tiaRoamingProfile_Session" & Fix(Rnd * 100000)
	Loop While fso.FolderExists(trpsmSessionDirectory)
	fso.CreateFolder trpsmSessionDirectory
End Sub

Function trpsmGetSessionDirectory()
	If trpsmSessionDirectory = "" Then
		Wscript.Echo "Error: This session is not initialized yet!"
		Wscript.Quit(1)
	End If
	trpsmGetSessionDirectory = trpsmSessionDirectory
End Function

Function trpsmPurgeSessionDirectory()
	If Not trpsmSessionDirectory = "" Then
		If fso.FolderExists(trpsmSessionDirectory) Then
			fso.DeleteFolder(trpsmSessionDirectory)
		End If
		trpsmSessionDirectory = ""
	End If
End Function
