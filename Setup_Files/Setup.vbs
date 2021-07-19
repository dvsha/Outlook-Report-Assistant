Dim checkExecutionPolicy, InstallDir, programspath
Set wshShell = CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
InstallDir = wshShell.ExpandEnvironmentStrings("%ProgramData%") & "\ORA"
programspath = wshShell.ExpandEnvironmentStrings("%ProgramData%") & "\Microsoft\Windows\Start Menu\Programs"

'-------------Checking and setting executionpolicy-----------------------------
checkExecutionPolicy = RegValueExists("HKLM","SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell","executionpolicy")
x641= x64()
If x641 Then
	if (NOT checkExecutionPolicy) then
		SetStringValue "HKLM", "SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell", "executionpolicy", "RemoteSigned"
		SetStringValue "HKLM", "SOFTWARE\Wow6432Node\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell", "executionpolicy", "RemoteSigned"
	end if
Else
	if (NOT checkExecutionPolicy) then
		SetStringValue "HKLM", "SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell", "executionpolicy", "RemoteSigned"
	end if
End If

'------------Copying Files-----------------
if NOT fso.FolderExists(InstallDir) then
	fso.CreateFolder (InstallDir)
end if
if NOT fso.FolderExists(InstallDir & "\File Repository") then
	fso.CreateFolder (InstallDir & "\File Repository")
end if

fso.CopyFile DirCurrent() & "\ORA.exe", InstallDir & "\" , True
fso.CopyFile DirCurrent() & "\ORA.exe.config", InstallDir & "\", True
fso.CopyFolder DirCurrent() & "\File Repository", InstallDir & "\", True

'-----------Creating Shortcut-------------------
if NOT fso.FolderExists(programspath & "\Outlook Report Assistant") then
	fso.CreateFolder(programspath & "\Outlook Report Assistant")
end if
Set link = wshShell.CreateShortcut(programspath & "\Outlook Report Assistant\ORA.lnk")
link.Arguments = ""
link.Description = "ORA shortcut"
link.IconLocation = InstallDir & "\File Repository\icon.ico"
link.TargetPath = InstallDir & "\ORA.exe"
link.WindowStyle = 3
link.WorkingDirectory = InstallDir
link.Save


Set fso = Nothing
Set wshShell = Nothing

'--------------Defining Functions---------------
Function RegValueExists(strHive,strKeyPath,strValueName)
	Dim strComputer, objRegistry, strValue
	Const HKCR = &H80000000
	Const HKCU = &H80000001
	Const HKLM = &H80000002
	Const HKU = &H80000003
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

	If strHive = "HKCR" then	
		objRegistry.GetStringValue HKCR,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKCU" then
		objRegistry.GetStringValue HKCU,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKLM" then
		objRegistry.GetStringValue HKLM,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKU" then
		objRegistry.GetStringValue HKU,strKeyPath,strValueName,strValue
	End If
	
	If IsNull(strValue) Then
		RegValueExists = False
	Else
		RegValueExists = True
	End If
	Set objRegistry = Nothing
End Function

Function SetStringValue(strHive,strKeyPath,strValueName, strValue)
	Dim strComputer1, objRegistry1
	Const HKCR = &H80000000
	Const HKCU = &H80000001
	Const HKLM = &H80000002
	Const HKU = &H80000003
	strComputer1 = "."
	Set objRegistry1 = GetObject("winmgmts:\\" & strComputer1 & "\root\default:StdRegProv")

	If strHive = "HKCR" then	
		objRegistry1.SetStringValue HKCR,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKCU" then
		objRegistry1.SetStringValue HKCU,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKLM" then
		objRegistry1.SetStringValue HKLM,strKeyPath,strValueName,strValue
	ElseIf strHive = "HKU" then
		objRegistry1.SetStringValue HKU,strKeyPath,strValueName,strValue
	End If

End Function


Function x64()
	Dim WshShell, OsType
	Set WshShell = CreateObject("WScript.Shell")
	OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
	If OsType = "x86" then
		x64 = False
	elseif OsType = "AMD64" then
		x64 = True
	end if
End Function

Public Function DirCurrent()
' Description: Returns directory of where script is running from.
	If Not BDEBUG Then On Error Resume Next
	Dim tCur, iSlash
	tCur = WScript.ScriptFullName
	iSlash = InStrRev(tCur, "\", -1, vbTextCompare)
	DirCurrent = Mid(tCur,1, iSlash-1)	
End Function
