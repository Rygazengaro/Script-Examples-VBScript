'This script sets everyone permissions on the nextgen folder for mass deployment
Function SetPermissions()
	Dim strNextgen, strHome
	Dim intRunError, objShell, objFSO
	strNextgen = "C:\NextGen"
	
	Set objShell = CreateObject("Wscript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	
	If objFSO.FolderExists(strNextgen) Then
		
		intRunError = objShell.Run("%COMSPEC% /c Echo Y| cacls " & strNextGen & " /t /e /c /p Everyone:F ", 2, True)
		
		If intRunError <> 0 Then
			Wscript.Echo "Error assigning permissions for user Everyone to Nextgen folder."
		End If
	End If
End Function

setPermissions()