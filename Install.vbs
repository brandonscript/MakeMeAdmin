On Error Resume Next
Const OverwriteExisting = True
Const ForWriting = 2
Dim objFSO, objShell, IsAdmin, objTextFile, strAddThese, Script, Task
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")
'Set IfUser = CreateObject("IsAdmin.Check")

USERNAME = WshShell.ExpandEnvironmentStrings("%USERNAME%")
PROGRAMFILES = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

SourceDirFind = Left(Wscript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
SourceDir = SourceDirFind
TargetDir = PROGRAMFILES & "\MakeMeAdmin"


'===========================AddThis.vbs SCRIPT============================
ScrAddThis ="'Dim additional variables as required." & VBCrLf &_
			"Dim strComp, objDstGrp, objSrcID" & VBCrLf &_
			"Const ForReading = 1" & VBCrLf &_
			"Set objFSO = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")"& VBCrLf &_
			"Set WshShell = WScript.CreateObject(" & chr(34) & "WScript.Shell" & chr(34) & ")" & VBCrLf &_
			VBCrLf &_
			"USERNAME = WshShell.ExpandEnvironmentStrings(" & chr(34) & "%USERNAME%" & chr(34) & ")" & VBCrLf &_
			"SYSTEMDRIVE = WshShell.ExpandEnvironmentStrings(" & chr(34) & "%SYSTEMDRIVE%" & chr(34) & ")" & VBCrLf &_
			VBCrLf &_
			"Set objTextFile = objFSO.OpenTextFile (SYSTEMDRIVE & " & chr(34) & "\Progra~1\MakeMeAdmin\ObjectList.olx" & chr(34) & ", ForReading)" & VBCrLf &_
			VBCrLf &_
			"'This specifies the local computer.  You could change this to a remote computer if you have admin rights on both the local AND remote computers." & VBCrLf &_
			"strComp = " & chr(34) & "." & chr(34) & VBCrLf &_
			VBCrLf &_
			"On Error Resume Next" & VBCrLf &_
			VBCrLf &_
			"Do Until objTextFile.AtEndOfStream" & VBCrLf &_
			chr(9) & "strLine = objTextFile.Readline" & VBCrLf &_
			chr(9) & "arrObjectList = Split(strLine," & chr(34) & ";" & chr(34) & ")" & VBCrLf &_
			chr(9) & "For i = 0 to Ubound(arrObjectList)" & VBCrLf &_
			chr(9) & chr(9) & "Set objDstGrp = GetObject(" & chr(34) & "WinNT://" & chr(34) & " & strComp & " & chr(34) & "/Administrators" & chr(34) & ") 'Destination local group" & VBCrLf &_
			'Configure this line, replacing FQDN
			chr(9) & chr(9) & "Set objSrcID = GetObject(" & chr(34) & "WinNT://{YOUR-FQDN}/" & chr(34) & " & arrObjectList(i)) 'Source domain account/group" & VBCrLf &_
			chr(9) & chr(9) & "objDstGrp.Add(objSrcID.ADsPath) 'Add account/group" & VBCrLf &_
			chr(9) & "Next" & VBCrLf &_
			"Loop" & VBCrLf
			

'========================MakeMeAdmin.xml XML TASK==========================

XMLMMA = 	"<?xml version=" & chr(34) & "1.0" & chr(34) & " encoding=" & chr(34) & "UTF-16" & chr(34) & "?>" & vbCrLf &_
			"<Task version=" & chr(34) & "1.2" & chr(34) & " xmlns=" & chr(34) & "http://schemas.microsoft.com/windows/2004/02/mit/task" & chr(34) & ">" & vbCrLf &_
			chr(9) & "<RegistrationInfo>" & vbCrLf &_
			chr(9) & chr(9) & "<Date>2015-05-07T13:35:01.5754822</Date>" & vbCrLf &_
			'Configure this line, replacing domain and author
			chr(9) & chr(9) & "<Author>{YOUR-DOMAIN}\{AUTHOR-ACCOUNT}</Author>" & vbCrLf &_
			chr(9) & "</RegistrationInfo>" & vbCrLf &_
			chr(9) & "<Triggers>" & vbCrLf &_
			chr(9) & chr(9) & "<BootTrigger>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<Repetition>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & chr(9) & "<Interval>PT1M</Interval>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & chr(9) & "<StopAtDurationEnd>false</StopAtDurationEnd>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "</Repetition>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<ExecutionTimeLimit>PT5M</ExecutionTimeLimit>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<Enabled>true</Enabled>" & vbCrLf &_
			chr(9) & chr(9) & "</BootTrigger>" & vbCrLf &_
			chr(9) & "</Triggers>" & vbCrLf &_
			chr(9) & "<Principals>" & vbCrLf &_
			chr(9) & chr(9) & "<Principal id=" & chr(34) & "Author" & chr(34) & ">" & vbCrLf &_
			chr(9) & chr(9) & "<UserId>SYSTEM</UserId>" & vbCrLf &_
			chr(9) & chr(9) & "<RunLevel>HighestAvailable</RunLevel>" & vbCrLf &_
			chr(9) & "</Principal>" & vbCrLf &_
			chr(9) & "</Principals>" & vbCrLf &_
			chr(9) & "<Settings>" & vbCrLf &_
			chr(9) & chr(9) & "<IdleSettings>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<Duration>PT10M</Duration>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<WaitTimeout>PT1H</WaitTimeout>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<StopOnIdleEnd>true</StopOnIdleEnd>" & vbCrLf &_
			chr(9) & chr(9) & chr(9) & "<RestartOnIdle>false</RestartOnIdle>" & vbCrLf &_
			chr(9) & chr(9) & "</IdleSettings>" & vbCrLf &_
			chr(9) & chr(9) & "<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>" & vbCrLf &_
			chr(9) & chr(9) & "<DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>" & vbCrLf &_
			chr(9) & chr(9) & "<StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>" & vbCrLf &_
			chr(9) & chr(9) & "<AllowHardTerminate>true</AllowHardTerminate>" & vbCrLf &_
			chr(9) & chr(9) & "<StartWhenAvailable>false</StartWhenAvailable>" & vbCrLf &_
			chr(9) & chr(9) & "<RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>" & vbCrLf &_
			chr(9) & chr(9) & "<AllowStartOnDemand>true</AllowStartOnDemand>" & vbCrLf &_
			chr(9) & chr(9) & "<Enabled>true</Enabled>" & vbCrLf &_
			chr(9) & chr(9) & "<Hidden>false</Hidden>" & vbCrLf &_
			chr(9) & chr(9) & "<RunOnlyIfIdle>false</RunOnlyIfIdle>" & vbCrLf &_
			chr(9) & chr(9) & "<WakeToRun>false</WakeToRun>" & vbCrLf &_
			chr(9) & chr(9) & "<ExecutionTimeLimit>P3D</ExecutionTimeLimit>" & vbCrLf &_
			chr(9) & chr(9) & "<Priority>7</Priority>" & vbCrLf &_
			chr(9) & "</Settings>" & vbCrLf &_
			chr(9) & "<Actions Context=" & chr(34) & "Author" & chr(34) & ">" & vbCrLf &_
			chr(9) & chr(9) & "<Exec>" & vbCrLf &_
			chr(9) & chr(9) & "<Command>" & chr(34) & "C:\Program Files\MakeMeAdmin\AddThis.vbs" & chr(34) & "</Command>" & vbCrLf &_
			chr(9) & chr(9) & "</Exec>" & vbCrLf &_
			chr(9) & "</Actions>" & vbCrLf &_
			"</Task>"

objFSO.CreateFolder("C:\PlaceHolder")
Return = WshShell.Run ("icacls C:\Placeholder /grant Users:(OI)(CI)M /T", 0, true)

If objFSO.FolderExists("C:\PlaceHolder") Then

	objFSO.CreateFolder(TargetDir)
	
	'Create the AddThis.vbs  file and fill it with the appropriate script
	Set objTextFile = objFSO.CreateTextFile(TargetDir & "\AddThis.vbs", 2, True)
	objTextFile.WriteLine(ScrAddThis)
	objTextFile.Close	
	
	'Create the MakeMeAdmin.xml file and fill it with the appropriate config
	Set objTextFile = objFSO.CreateTextFile(TargetDir & "\MakeMeAdmin.xml", 2, True)
	objTextFile.WriteLine(XMLMMA)
	objTextFile.Close	
	
	'These are no longer needed as the files are being created instead of copied.
	'objFSO.CopyFile SourceDir & "AddThis.vbs",TargetDir,OverwriteExisting
	'objFSO.CopyFile SourceDir & "MakeMeAdmin.xml",TargetDir,OverwriteExisting
			
	strAddThese = inputbox("Enter the domain objects you wish to add to the Administrators group, separating each with a semicolon (;), i.e.: " _
		& vbCrLf & vbCrLf & "     tomjones;some-active-directory-group;amypetersson","Choose Objects",USERNAME)
	
	If strAddThese = "" Then
		objFSO.DeleteFolder("C:\PlaceHolder")
		objFSO.DeleteFolder(TargetDir)
		wscript.quit
	End If
	
	Set objTextFile = objFSO.OpenTextFile(TargetDir & "\ObjectList.olx", 2, True)
	objTextFile.WriteLine(strAddThese)
	objTextFile.Close
		
	Return = WshShell.Run ("schtasks /create /TN MakeMeAdmin /XML " & chr(34) & TargetDir & "\MakeMeAdmin.xml" & chr(34), 0, true)
	wscript.sleep 4000
	Return = WshShell.Run ("schtasks /run /TN MakeMeAdmin", 0, true)

	If msgbox("The specified object(s) should now have been added to the Administrators group. Do you want to open Computer Management to check?",VBYesNo,"Open Computer Management?") = vbYes Then
		WshShell.Run SYSTEMDRIVE & "\Windows\System32\mmc.exe compmgmt.msc"
	End If
	
	objFSO.DeleteFolder("C:\PlaceHolder")
	'wscript.echo "icacls " & chr(34) & TargetDir & chr(34) & " /grant Users:(OI)(CI)M /T"
	Return = WshShell.Run ("icacls " & chr(34) & TargetDir & chr(34) & " /grant Users:(OI)(CI)M /T", 0, true)
Else
	Elevate
End If	

Sub Elevate()
	objShell.ShellExecute WScript.FullName, chr(34) & WScript.ScriptFullName & chr(34), vbNullString, "runas"
End Sub

