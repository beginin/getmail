'==== Script Information Header ====
'script name:  Purge Temp
'version:		1.0
'date:			16.07.08
'autor:			Beginin Vladimir
'description:	Скрипт удаляет устаревшие временные.

'==== Script Main Logic ====
on error resume next
const PurgeTime = 30 'days

'Exceptions - folders, which will not be processed
dim aExceptions(3)
aExceptions(0) = "Default User"
aExceptions(1) = "LocalService"
aExceptions(2) = "NetworkService"
'aExceptions(3) = "All Users"

set oFSO = CreateObject("Scripting.Filesystemobject")
set oShell = createobject("wscript.shell")

'Set paths
sProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
sWinDir = oShell.ExpandEnvironmentStrings("%WinDir%")
sWinTempFolder = sWinDir & "\Temp"
sDocuments = "C:\Documents and Settings"
sTest="E:\Getmailinbox\"
sLogFileName="E:\Getmailinbox\log\Purge_"

'Create log-file
sLogFileName = sLogFileName & Date 
Set oLogFile = oFSO.CreateTextFile(sLogFileName & ".log", true)
oLogFile.WriteLine "========== Start purging =========="
 
'Purge Test folder
oLogFile.WriteLine vbCrLf & "========== Test folder =========="
PurgeFolder(sTest)

'Close log-file
oLogFile.WriteLine vbCrLf & "========== Stop purging =========="
oLogFile.Close


Set objMsg = CreateObject("CDO.Message") 
Set Config = CreateObject("CDO.Configuration") 
Set Config = objMsg.Configuration 
objMsg.From = "Purge@add.astron.local" 
objMsg.To = "admins@post.flagman.int" 
objMsg.Subject = "Очистка произведена" 
objMsg.AddAttachment(sLogFileName & ".log")
objMsg.Textbody = "Лог файлов приклеплен" 
Config("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Config("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "post.flagman.int" 
Config("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
Config("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
Config.Fields.Update 
objMsg.Send 

'PurgeFolder procedure
sub PurgeFolder(sFolderPath)
	set oFolder = oFSO.GetFolder(sFolderPath)
	set colFiles = oFolder.Files
	for each oFile in colFiles
		if (Date-oFile.DateLastModified) > PurgeTime and (Date-oFile.DateCreated) > PurgeTime then
			oLogFile.Writeline oFile.Path & vbTab & oFile.DateCreated
			oFSO.DeleteFile oFile.Path, true
			if err.Number <> 0 then
				oLogFile.Writeline "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				err.clear
			end if
			wscript.sleep 20
		end if
	next
	set colSubFolders = oFolder.SubFolders
	for each oSubFolder in colSubFolders
		PurgeFolder(oSubFolder.Path)
		if oSubFolder.Size = 0 then 
			oLogFile.Writeline oSubFolder.Path & vbTab & oSubFolder.DateCreated
			oFSO.DeleteFolder oSubFolder.path
			if err.Number <> 0 then
				oLogFile.Writeline "-----> Error # " & CStr(Err.Number) & " " & Err.Description
				err.clear
			end if
		end if
	next
end sub
