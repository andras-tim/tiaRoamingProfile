' tiaRoamingProfile v0.1.131014
' Created by Andras Tim @ 2013
'
' TODO:
'   * dynamic load reg files and others from preconf dir
'   * if not exists preconf dir then have to create

''' MODULES '''


''' MAIN '''
Sub trppcPrepare()
	trppcPrepareRegistry
End Sub

Sub trppcPrepareRegistry()
	wshShell.Run "%ComSpec% /c reg import """ & prgPath & "\preconf\console.reg""", 0, True
End Sub
