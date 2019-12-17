' VBScript Template
' Author: Tomas Kadlec, info@tomkadlec.cz
' Version 1.0
' Vyplnit radek 63 promena "sPackageName"
'----------------------------------------
' AUTOR SCRIPTU :
' DATUM         :
' POPIS         :
'----------------------------------------
'==============================================
' NASTAVENI PROMENYCH, NEUPRAVOVAT
'==============================================
Option Explicit
Dim oShell, oFSO
Dim LogPath, LogFile, oLogFile, LogText
Dim RunCommand,RunParameter
Dim LastLoggedOnUser, UserProfile, UserAppData, UserStartMenu, UserDesktop
Dim AllUsersProfile, AllUsersDesktop, AllUsersStartMenu
Dim ProgramFiles, ProgramFilesX86, Temp
Dim ScriptPath, SystemDrive, WinDir, WinDirSys32
Dim MessageText, FolderName, FileName
Dim sPackageName
Dim RegExistsValue, sRegExistsValue
Dim sRegGetValue, RegGetValueName
Dim RegSetValueName, RegSetValue, RegSetValueType
Dim DelFileName
Dim MsiFileName, MsiInstParameter, MsiDeInstallGUID, MsiDeInstallParameter
Dim AppName

Set oShell = WScript.CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

LastLoggedOnUser	= REGGET("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\LastLoggedOnUser")
If InStrRev(LastLoggedOnUser, "\") Then
	LastLoggedOnUser = Right(LastLoggedOnUser, InStrRev(LastLoggedOnUser, "\"))
End If
LogPath				= SystemDrive & "\applogs"												'c:\applogs
AllUsersProfile     = oShell.ExpandEnvironmentStrings("%AllUsersProfile%") & "\"			'c:\programData\
AllUsersDesktop     = oShell.SpecialFolders ("AllUsersDesktop") & "\"						'c:\users\public\Desktop\
AllUsersStartMenu   = oShell.SpecialFolders ("AllUsersStartMenu") & "\"						'c:\programData\microsoft\windows\Start Menu\
Temp                = oShell.ExpandEnvironmentStrings("%Temp%") & "\"						'c:\Users\%Username%\AppData\local\temp OR c:\windows\temp (SystemUser)
ProgramFiles        = oShell.ExpandEnvironmentStrings("%ProgramFiles%") & "\"		 		'c:\Program Files\
ProgramFilesX86 	= oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\"			'c:\Program Files (x86)\
ScriptPath          = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))	'Path of this Script with '\' of the End
SystemDrive         = oShell.ExpandEnvironmentStrings("%SystemDrive%") & "\"				'c:\
WinDir              = oShell.ExpandEnvironmentStrings("%WinDir%") & "\"						'c:\windows\
WinDirSys32			= WinDir & "system32\"													'c:\windows\system32\
UserProfile			= SystemDrive & "Users\" & LastLoggedOnUser & "\"						'c:\Users\%Username%\
UserAppData			= UserProfile & "AppData\Roaming\"										'c:\Users\%UserName%\AppData\Roaming\
UserStartMenu		= UserAppData & "Microsoft\Windows\Start Menu\"							'c:\Users\%UserName%\AppData\Roaming\Microsoft\Windows\Start Menu\
UserDesktop			= UserProfile & "Desktop\"												'c:\Users\%UserName%\Desktop\

Const HKCR = &H80000000		'HKEY_CLASSES_ROOT
Const HKCU = &H80000001		'HKEY_CURRENT_USER
Const HKLM = &H80000002		'HKEY_LOCAL_MACHINE
Const HKUS = &H80000003		'HKEY_USERS
Const HKCC = &H80000005		'HKEY_CURRENT_CONFIG

'==============================================
' NAZEV BALICKU
'==============================================

sPackageName = ""

'==============================================
' KONFIGURACE LOGOVANI
'==============================================
If sPackageName = "" Then
	LogFile	 			= LogPath & "\" & mid(Mid(ScriptPath,1,Len(ScriptPath)-1),InStrRev(Mid(ScriptPath,1,Len(ScriptPath)-1),"\")+1) & "_" & oFSO.GetBaseName (WScript.ScriptName) & ".log"
Else
	LogFile	 			= LogPath & "\" & sPackageName & "_" & oFSO.GetBaseName (WScript.ScriptName) & ".log"
End If
If oFSO.FileExists(LogFile) Then
	On Error Resume Next
	oFSO.DeleteFile LogFile
	On Error GoTo 0
End If
LOG ""
LOG "START of the script : " & ScriptPath & oFSO.GetBaseName (WScript.ScriptName) & ".vbs"
LOG ""

'==============================================
' SCRIPT PROMENE
'==============================================



'==============================================
' SCRIPT KOD
'==============================================






'==============================================
' SCRIPT KONEC
'==============================================
On Error Resume Next
LOG ""
LOG "END of the script : " & scriptpath & oFSO.GetBaseName (WScript.ScriptName) & ".vbs"
Set oShell = Nothing
Set oFSO = Nothing
WScript.Quit


'==============================================
' FUNKCE
'==============================================

' Function create log file
Function LOG (LogText)
	On Error Resume Next
	If Not oFSO.FolderExists(LogPath) Then
		oFSO.CreateFolder LogPath
	End If
	Set oLogFile = oFSO.OpenTextFile(LogFile,8,True)
	If instr(Date," ") <> 0 Then
		oLogFile.WriteLine trim(Date) & " " & Time & "  :  " & LogText & vbCr
	Else
		oLogFile.WriteLine Date & " " & Time & "  :  " & LogText & vbCr
	End If
	oLogFile.Close
End Function

' Function Run command
Function RUN (RunCommand,RunParameter)
	If oFSO.FileExists(RunCommand) Then
		LOG "Function RUN :" & RunCommand & " " & Chr(34) & RunParameter & Chr(34)
		oShell.Run Chr(34) & RunCommand & Chr(34) & " " & RunParameter,1,True
	Else
		LOG "Cannot RUN " & Chr(34) & RunCommand & Chr(34) & " ==> File not exists"
	End If
End Function

' Function show messagebox
Function MESSAGE (MessageText)
	LOG "Function MESSAGE: " & Chr(34) & MessageText & Chr(34)
	MsgBox MessageText
End Function

' Function create folder
Function CREATEFOLDER(FolderName)
	If oFSO.FolderExists(FolderName) Then
		LOG "Function CREATEFOLDER: Folder " & Chr(34) & FolderName & Chr(34) & " already exists"
	Else
		oFSO.CreateFolder FolderName
		LOG "Function CREATEFOLDER: Folder " & Chr(34) & FolderName & Chr(34) & " created"
	End If
End Function

' Function delete folder
Function DELFOLDER(FolderName)
	If oFSO.FolderExists(FolderName) Then
		oFSO.DeleteFolder FolderName
		LOG "Function DELFOLDER: Folder " & Chr(34) & FolderName & Chr(34) & " deleted"
	Else
		LOG "Function DELFOLDER: Folder " & Chr(34) & FolderName & Chr(34) & " not exists"
	End If
End Function

' Function copy file
Function COPYFILE(FileName,FolderName)
	If oFSO.FileExists(FileName) Then
		If oFSO.FolderExists(FolderName) Then
			oFSO.CopyFile FileName,FolderName,True
			LOG "Function COPYFILE: File: " & Chr(34) & FileName & " copied to folder " & FolderName & Chr(34)
		Else
			LOG "Function COPYFILE: Destination Folder: " & Chr(34) & FolderName & Chr(34) & " not exists"
		End If
	Else
		LOG "Function COPYFILE: File: " & Chr(34) & FileName & " not exists!"
	End If

End Function

' Function delete file
Function DELFILE(DelFileName)
	If oFSO.FileExists(DelFileName) Then
		oFSO.DeleteFile(DelFileName)
		LOG "Function DELFILE: File " & Chr(34) & DelFileName & Chr(34) & " deleted"
	Else
		LOG "Function DELFILE: File " & Chr(34) & DelFileName & Chr(34) & " not exists!"
	End If
End Function

' Function check if value exists in Registry
Function REGEXISTS(RegExistsValue)
	On Error Resume next
	sRegExistsValue = oShell.RegRead (RegExistsValue)
	REGEXISTS = (Err.Number = 0)
	on Error goto 0
	if REGEXISTS = true Then
		LOG ("Function REGEXISTS: The ValueName: " & RegExistsValue & " is Exists!")
	Else
		LOG ("Function REGEXISTS: The ValueName: " & RegExistsValue & " is NOT Exists!")
	End If
End Function

' Function get value from Registry
Function REGGET(RegGetValueName)
	If REGEXISTS (RegGetValueName) = True Then
		sRegGetValue = oShell.RegRead (RegGetValueName)
	Else
		sRegGetValue = ""
	End If
	LOG ("Function REGGET: " & "The Content of: " & RegGetValueName & " is: '" & sRegGetValue & "'")
	REGGET = sRegGetValue
End Function

' Function set value to Registry
Function REGSET(RegSetValueName,RegSetValue,RegSetValueType)
	If REGEXISTS(RegSetValueName) = True Then
		oShell.RegWrite RegSetValueName,RegSetValue,RegValueType
		LOG "Function REGSET: Reg Value Name: " & RegSetValueName & " set to: " & RegSetValue & " value type: " & RegSetValueType
	End If
End Function

' Function delete value from Registry
Function REGDEL(RegDelName)
	If REGEXISTS(RegDelName) = True Then
		oShell.RegDelete RegDelName
		LOG ("Function REGDEL: Reg key: ") & RegDelName & " deleted!"
	End If
End Function

' Function install MSI file
Function MSIINST(MsiFileName,MsiInstParameter)
	If oFSO.FileExists(MsiFileName) = True Then
		If MsiInstParameter = "" Then
			RUN WinDirSys32 & "msiexec.exe", "/I " & Chr(34) & MsiFileName & Chr(34) & " /QB!"
			LOG "Function MSIINST: Install MSI file: " & MsiFileName
		Else
			RUN WinDirSys32 & "msiexec.exe", "/I " & Chr(34) & MsiFileName & Chr(34) & " " & MsiInstParameter
			LOG "Function MSIINST: Install MSI file: " & MsiFileName & " MSI Parameter: " & MsiInstParameter
		End If
	Else
		LOG "Function MSIINST: MSI File Not Found! " & MsiFileName
	End If
End Function

' Function uninstall MSI file
Function MSIDEINST(MsiDeInstallGUID,MsiDeInstallParameter)
	If REGEXISTS("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & MsiDeInstallGUID & "\DisplayName") = True Then
		If MsiDeInstallParameter = "" Then
			RUN WinDirSys32 & "msiexec.exe", "/x " & MsiDeInstallGUID & " /QB!"
			LOG "Function MSIDEINST: MSI GUID: " & MsiDeInstallGUID & " Deinstalled!"
		Else
			RUN WinDirSys32 & "msiexec.exe", "/x " & MsiDeInstallGUID & " " & MsiDeInstallParameter
			LOG "Function MSIDEINST: MSI GUID: " & MsiDeInstallGUID & " with parameter " & MsiDeInstallParameter & " Deinstalled!"
		End If
	Else
		If REGEXISTS("HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & MsiDeInstallGUID & "\DisplayName") = True Then
			If MsiDeInstallParameter = "" Then
				RUN WinDirSys32 & "msiexec.exe", "/x " & MsiDeInstallGUID & " /QB!"
				LOG "Function MSIDEINST: MSI GUID: " & MsiDeInstallGUID & " Deinstalled!"
			Else
				RUN WinDirSys32 & "msiexec.exe", "/x " & MsiDeInstallGUID & " " & MsiDeInstallParameter
				LOG "Function MSIDEINST: MSI GUID: " & MsiDeInstallGUID & " with parameter " & MsiDeInstallParameter & " Deinstalled!"
			End If
		Else
			LOG "Function MSIDEINST: MSI GUID: " & MsiDeInstallGUID & " not found!"
		End If
	End If

End Function

' Function reboot machine
Function REBOOT
	MsgBox "Po kliknuti na OK bude Vas system restartovan!"
	LOG "Function REBOOT"
	oShell.Run "shutdown -r -f -t 0",1,True
End Function

'
