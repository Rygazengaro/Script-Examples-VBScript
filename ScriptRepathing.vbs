' This replaces the network share locations in all the scripts after a network share migration
Set oShell = CreateObject("WScript.Shell")
Set obj = CreateObject("Scripting.FileSystemObject")
Dim objWMIService, colItems, objItem, strComputer, objFile, objFolder, System, Shell, ShellCMD, Params, objNetwork, colProcessList
strComputer = "."
objStartFolder = "\\networkshare2"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objFolder = obj.GetFolder(objStartFolder)
Const ForReading = 1
Const ForWriting = 2

'***************************************************
' Find all script files, and replace incorrect paths
'***************************************************
Call SubfoldersPS1(obj.GetFolder(objStartFolder))
Call SubfoldersBAT(obj.GetFolder(objStartFolder))
Call SubfoldersVBS(obj.GetFolder(objStartFolder))

WScript.Quit(0)

'***************** SUB PROCEDURES ****************|
'---------------PLEASE DO NOT CHANGE--------------|
'*************************************************|

Sub SubFoldersPS1(Folder)
	For Each Subfolder in Folder.SubFolders
		Set objFolder = obj.GetFolder(Subfolder.Path)
		Set colFiles = objFolder.Files
		For Each objItem in colFiles
			if LCase(obj.GetExtensionName(objItem.Name)) = "ps1" then
				Set objFile = obj.OpenTextFile(objItem, ForReading)
					strText = objFile.ReadAll
					objFile.Close
					strNewText = Replace(strText, "\\networkshare1", "\\networkshare2")
				Set objFile = obj.OpenTextFile(objItem, ForWriting)
					objFile.WriteLine(strNewText)
					objFile.Close			
			end if
		Next
		SubFoldersPS1 Subfolder
	Next
End Sub

Sub SubFoldersVBS(Folder)
	For Each Subfolder in Folder.SubFolders
		Set objFolder = obj.GetFolder(Subfolder.Path)
		Set colFiles = objFolder.Files
		For Each objItem in colFiles
			if LCase(obj.GetExtensionName(objItem.Name)) = "vbs" then
				Set objFile = obj.OpenTextFile(objItem, ForReading)
					strText = objFile.ReadAll
					objFile.Close
					strNewText = Replace(strText, "\\networkshare1", "\\networkshare2")
				Set objFile = obj.OpenTextFile(objItem, ForWriting)
					objFile.WriteLine(strNewText)
					objFile.Close			
			end if
		Next
		SubFoldersVBS Subfolder
	Next
End Sub

Sub SubFoldersBAT(Folder)
	For Each Subfolder in Folder.SubFolders
		Set objFolder = obj.GetFolder(Subfolder.Path)
		Set colFiles = objFolder.Files
		For Each objItem in colFiles
			if LCase(obj.GetExtensionName(objItem.Name)) = "bat" then
				Set objFile = obj.OpenTextFile(objItem, ForReading)
					strText = objFile.ReadAll
					objFile.Close
					strNewText = Replace(strText, "\\networkshare1", "\\networkshare2")
				Set objFile = obj.OpenTextFile(objItem, ForWriting)
					objFile.WriteLine(strNewText)
					objFile.Close			
			end if
		Next
		SubFoldersBAT Subfolder
	Next
End Sub

