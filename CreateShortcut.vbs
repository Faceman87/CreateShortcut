'First line does not nee to be repeated for multiple files/folders
Set oWS = WScript.CreateObject("WScript.Shell")

'Create Shortcut for name of desired FILE. include .LNK at end of name
'.LNK is a shorcut type
sLinkFile = "C:\Users\me\Desktop\SAMPLEFILE.LNK"

Set oLink = oWS.CreateShortcut(sLinkFile)
	oLink.TargetPath = "C:\Users\me\Documents and Settings\SAMPLEFILE.xlsx"
	
oLink.Save

'Create Shortcut for name of desired FOLDER. include .LNK at end of name
sLinkFile = "C:\Users\me\Desktop\SAMPLEFOLDER.LNK"

Set oLink = oWS.CreateShortcut(sLinkFile)
	oLink.TargetPath = "C:\Users\me\Documents and Settings\SAMPLEFOLDER"
	
oLink.Save
