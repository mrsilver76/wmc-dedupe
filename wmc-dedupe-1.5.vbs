Option Explicit

'
' wmc-dedupe (version 1.5, 23rd October 2014)
' Copyright © 2011-2014 Richard Lawrence
' https://github.com/mrsilver76/wmc-dedupe
'
' A program which identifies duplicate Windows Media Center recorded
' television shows (in either WTV or DVR-MS formats) and then either
' moves them into a folder for duplicates or deletes them. Recordings
' that are sitting in the duplicates folder can be automatically deleted
' after a certain number of days.
'
' This program is free software; you can redistribute it and/or modify it
' under the terms of the GNU General Public License as published by the
' Free Software Foundation; either version 2 of the License, or (at your
' option) any later version.
'
' This program is distributed in the hope that it will be useful, but
' WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
' Public License for more details.
'
' ========================================================================
'

' Some defaults

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim oShell : Set oShell = CreateObject("Shell.Application")

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const VERSION = "1.5"

' Settings
Dim bForceRun, bVerboseLog, bDelete, bSilent, bTest, bPauseBeforeQuit, bNoDVRMS, bNoWTV, bVeryVerboseLog, iEraseDays, bMove, bErase
bForceRun = False : bVerboseLog = False : bDelete = False : bSilent = False : bTest = False : bPauseBeforeQuit = False : bMove = False
bErase = False : bNoDVRMS = False : bNoWTV = False : bVeryVerboseLog = False : iEraseDays = -1
' Paths to files, logs, config and the lock file
Dim sRecordedTV, sLogName, sLockFile, sDuplicateFolder
' TV show meta-data
Dim oFolderA, oFolderB
Dim	sTVTitle, sTVSubtitle, sTVDescription
' File details
ReDim sFilename(-1), sFilePath(-1), iFileSize(-1), sFileTitle(-1), sFileSubTitle(-1), sFileDesc(-1), bTestMoved(-1)
' Statistics
Dim iMoved, iDeleted, iTotal : iMoved = 0 : iDeleted = 0 : iTotal = 0

' Main code

Call Read_Params
Call Force_Cscript_Execution
Call Prepare_Logging
Call Check_Lock_File
Call Check_Not_Recording
Call Scan_Files
Call Erase_Duplicates
Call Delete_Lock_File
Call Shutdown

' Read_Params
' Work out what command line arguments (if any) have been passed to this
' program and set the global flags relating to them.

Sub Read_Params

	Dim iCount
	Dim sPath(2), sRecTV : sRecTV = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%PUBLIC%") & "\Recorded TV"
	Dim bUseRecordedPath : bUseRecordedPath = False
	sRecordedTV = ""
	
	For iCount = 0 to WScript.Arguments.Count - 1
		Select Case LCase(WScript.Arguments(iCount))
			Case "/?", "-h", "--help"
				Call Display_Usage("")
			Case "/t", "-t", "--test"
				' Test mode, files are not actually deleted
				bTest = True
			Case "/p", "-p", "--pause"
				' Pause (and display "press [ENTER]" prompt) before quitting
				bPauseBeforeQuit = True
			Case "/f", "-f", "--force"
				' Force this script to run even if something is being recorded
				bForceRun = True
			Case "/v", "-v", "--verbose"
				' Verbose logging
				bVerboseLog = True
			Case "/vv", "-vv", "--veryverbose"
				' Very Verbose logging
				bVerboseLog = True
				bVeryVerboseLog = True
			Case "/s", "-s", "--silent"
				' Don't display anything to the screen
				bSilent = True
			Case "/r", "-r", "--recordedtv"
				' Use the Public Recorded TV path
				bUseRecordedPath = True
			Case "/d", "-d", "--delete"
				' Delete shows that are duplicates
				bDelete = True
			Case "/m", "-m", "--move"
				' Move shows that are duplicates
				bMove = True
			Case "/nd", "-nd", "--nodvrms"
				' Don't examine DVR-MS files
				bNoDVRMS = True
			Case "/nw", "-nw", "--nowtv"
				' Don't examine WTV files
				bNoWTV = True
			Case Else
				If Left(WScript.Arguments(iCount), 1) = "/" or Left(WScript.Arguments(iCount), 1) = "-" Then
					If Parse_Erase(LCase(WScript.Arguments(iCount))) = False Then
						' This is an unknown argument
						Call Display_Usage("Unknown option (" & WScript.Arguments(iCount) & ")")
					End If
				Else
					' This is a path to something
					If sPath(1) = "" Then
						sPath(1) = WScript.Arguments(iCount)
					Else
						sPath(2) = WScript.Arguments(iCount)
					End If
				End If
		End Select
	Next 
	
	
	If bUseRecordedPath = True Then
		sRecordedTV = sRecTV
		sDuplicateFolder = sPath(1)
	ElseIf bMove = True Or bDelete = True Then
		sRecordedTV = sPath(1)
		sDuplicateFolder = sPath(2)
	ElseIf bErase = True Then 
		sRecordedTV = ""
		sDuplicateFolder = sPath(1)
	Else
		Call Display_Usage("Missing /M, /D or /E:x")
	End If
		
	' Check the right locations have been provided
	
	If bMove = True and bDelete = True Then Call Display_Usage("Either /M or /D should be used, but not both")
	
	If bMove = True or bDelete = True Then
		If sRecordedTV = "" Then Call Display_Usage("Path to recorded television files (or /R) is missing")
		If fso.FolderExists(sRecordedTV) = False Then Call Display_Usage("Supplied recorded TV path does not exist (" & sRecordedTV & ")")
	End If
	
	If bMove = True Or bErase = True Then
		If sDuplicateFolder = "" Then Call Display_Usage("Path to the Duplicates folder is missing")
		If fso.FolderExists(sDuplicateFolder) = False Then Call Display_Usage("Supplied Duplicate path does not exist (" & sDuplicateFolder & ")")
		If LCase(sDuplicateFolder) = LCase(sRecTV) Then Display_Usage("Duplicate path is same as Public 'Recorded TV' path. Not recommended")
		If LCase(sDuplicateFolder) = LCase(sRecordedTV) Then Display_Usage("Duplicate and Recorded TV paths are the same. Not recommended	")
	End If
			
End Sub

' Parse_Erase
' Given an argument, workout what the EraseDays value is.

Function Parse_Erase(sArg)

	Dim iValue
	Parse_Erase = False

	If Len(sArg) < 4 Then
		If Left(sArg, 2) = "/e" or Left(sArg, 2) = "-e" Then
			Call Display_Usage("Missing value (days) with option /E")
		End If
		Exit Function 
	End If
	
	If Left(sArg, 3) = "/e:" Or Left(sArg, 3) = "-e=" Then
		On Error Resume Next
		iValue = CInt(Mid(sArg, 4, Len(sArg)))
		On Error Goto 0
	ElseIf Len(sArg) > 8 And Left(sArg, 8) = "--erase=" Then
		On Error Resume Next
		iValue = CInt(Mid(sArg, 9, Len(sArg)))
		On Error Goto 0
	Else
		' To do: Flag if the Erase value is badly formed
		Exit Function
	End If
	
	If iValue < 1 Then
		Call Display_Usage("Erase value must be 1 or greater")
		Exit Function
	End If
	
	iEraseDays = iValue
	bErase = True
	Parse_Erase = True
	
End Function

' Display_Usage
' Explain to the user how the script works. If sError contains a string
' then this is appended to the bottom.

Sub Display_Usage(sError)

	Dim sText
	Dim sName : sName = Wscript.ScriptName
	
	If Instr(sName, " ") <> 0 Then
		sName = """" & sName & """"
	End If
	
	sText = "wmc-dedupe (version " & VERSION & ")" & VbCrLf
	sText = sText & "Copyright © 2011-2014 Richard Lawrence." & VbCrLf
	sText = sText & "https://github.com/mrsilver76/wmc-dedupe" & VbCrLf
	sText = sText & VbCrLf
	sText = sText & "Usage: " & sName & " [/M | /D] [/R | <tv path>] [/E:x] <dup path> [/T] [/F] [/S] [/NW|/ND] [/P] [/V | /VV] " & VbCrLf
	sText = sText & VbCrLf
	sText = sText & "    No args     Display help. This is the same as typing /?." & VbCrLf
	sText = sText & "    /?          Display help. This is the same as not typing any options." & VbCrLf
	sText = sText & "    /M          Move duplicate recordings into another folder." & VbCrLf
	sText = sText & "    /D          Delete duplicate recordings." & VbCrLf
	sText = sText & "    /R          Look at shows in the Public 'Recorded TV' location." & VbCrLf
	sText = sText & "    <tv path>   The path to Recorded TV files. Required unless /R is used." & VbCrLf
	sText = sText & "    /E:x        Erase files in the duplicates folder older than x days." & VbCrLf
	sText = sText & "    [dup path]  Path to duplicates folder. Required with /M or /E." & VbCrLf
	sText = sText & "    /T          Test mode. Don't move or delete any shows." & VbCrLf
	sText = sText & "    /F          Force execution even if something is being recorded." & VbCrLf	
	sText = sText & "    /S          Silent. Don't display anything on the screen." & VbCrLf
	sText = sText & "    /NW         Ignore WTV files." & VbCrLf
	sText = sText & "    /ND         Ignore DVR-MS files." & VbCrLf
	sText = sText & "    /P          Pause after running." & VbCrLf
	sText = sText & "    /V          Verbose mode. Log additional information during execution." & VbCrLf
	sText = sText & "    /VV         Very Verbose mode. Log even more information. Implies /V."

	If sError <> "" Then
		sText = sText & VbCrLf & VbCrLf & "Error: " & sError
	End If
	
	WScript.Echo sText
	
	' Don't ask for confirmation to exit because this will appear as a second popup
	bPauseBeforeQuit = False
	Call Shutdown
	
End Sub

' Shutdown
' Terminate the script, displaying a confirmation if needs be. 

Sub Shutdown

	Call Log("wmc-dedupe stopped", False)
	Call Log("", False)
	
	Set fso = Nothing
	Set oShell = Nothing

	If bPauseBeforeQuit = True And bSilent = False Then	
		Wscript.Echo
		Wscript.Echo "Press [ENTER] or [RETURN] to end this program."
		Do While Not Wscript.StdIn.AtEndOfLine
			Dim sInput : sInput = Wscript.StdIn.Read(1)
		Loop
	End If
	
	WScript.Quit

End Sub

' Force_Cscript_Execution
' Force the script to run using cscript instead of wscript.

Sub Force_Cscript_Execution

	If bSilent = True Then Exit Sub

	Dim Arg, Str
    If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then
        For Each Arg In WScript.Arguments
			If InStr(Arg, " ") Then Arg = """" & Arg & """"
			Str = Str & " " & Arg
		Next
        CreateObject("WScript.Shell" ).Run "cscript //nologo """ & WScript.ScriptFullName & """ " & Str
        WScript.Quit
    End If
	
End Sub

' Prepare_Logging
' Deletes any old log files and prepares the new one for writing to

Sub Prepare_Logging
	
	' Create the appropriate folders (Application Data = &H1A&)
	Dim sApplicationData : sApplicationData = oShell.Namespace(&H1A&).Self.Path & "\wmc-dedupe"
	If fso.FolderExists(sApplicationData) = False Then fso.CreateFolder(sApplicationData)
	sLockFile = sApplicationData & "\lockfile.txt"

	' Create (if required) the folder containing the logs
	sApplicationData = sApplicationData & "\Logs"
	If fso.FolderExists(sApplicationData) = False Then fso.CreateFolder(sApplicationData)

	' Determine the name of the logfile	and lockfile
	sLogName = sApplicationData & "\Log " & Replace(FormatDateTime(Now(), 2), "/", "-") & ".txt"
	
	If bSilent = False Then
		Wscript.Echo "wmc-dedupe (version " & VERSION & ")"
		Wscript.Echo "Copyright © 2011-2014 Richard Lawrence <richard@fourteenminutes.com>"
		Wscript.Echo "http://www.fourteenminutes.com/code/wmc-dedupe/"
		Wscript.Echo
		Wscript.Echo "A program which identified duplicate Windows Media Center recorded"
		Wscript.Echo "television shows (in either WTV or DVR-MS formats) and then either"
		Wscript.Echo "moves them into a folder for duplicates or deletes them. Recordings"
		Wscript.Echo "that are sitting in the duplicates folder can be automatically deleted"
		Wscript.Echo "after a certain number of days."
		Wscript.Echo
		Wscript.Echo "This program is free software; you can redistribute it and/or modify it"
		Wscript.Echo "under the terms of the GNU General Public License as published by the"
		Wscript.Echo "Free Software Foundation; either version 2 of the License, or (at your"
		Wscript.Echo "option) any later version."
		Wscript.Echo
		Wscript.Echo "This program is distributed in the hope that it will be useful, but"
		Wscript.Echo "WITHOUT ANY WARRANTY; without even the implied warranty of"
		Wscript.Echo "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General"
		Wscript.Echo "Public License for more details. "
		Wscript.Echo
	End If

	Call Log("wmc-dedupe (version " & VERSION & ") starting", False)
		
	' Log variables used
	
	Call Log("Configuration: ForceRun = " & bForceRun & ", Verbose = " & bVerboseLog, True)
	Call Log("Configuration: Silent = " & bSilent & ", Test = " & bTest, True)
	Call Log("Configuration: Delete = " & bDelete & ", Move = " & bMove, True)
	Call Log("Configuration: Pause = " & bPauseBeforeQuit & ", Erase = " & bErase, True)
	Call Log("Configuration: NoDVRMS = " & bNoDVRMS & ", NoWTV = " & bNoWTV, True)
	Call Log("Configuration: VeryVerbose = " & bVeryVerboseLog & ", EraseDays = " & iEraseDays, True)
	If bNoWTV = True And bNoDVRMS = True Then
		Call Log("Warning: Both /NW and /ND flags are being used. Nothing will be examined!", False)
	End IF
	
	Call Log("Configuration: RecordedTV = " & sRecordedTV, True)
	Call Log("Configuration: Duplicates = " & sDuplicateFolder, True)
	
	' Clean up log files
	
	Dim oFiles : Set oFiles = fso.GetFolder(sApplicationData).Files
	Dim oFile
	For Each oFile In oFiles
	
		' If the log file has not been modified in over 7 days, then
		' delete it.
		
		Dim iAge : iAge = DateDiff("d", oFile.DateLastModified, Now()) 
		Call Log("Examining log file: " & oFile.Name & " (age: " & iAge & " days)", True)
	
		If Left(oFile.Name, 4) = "Log " And iAge > 7 Then
			' Delete this file
			Call Log("Deleting old log file: " & oFile.Name, True)
			fso.DeleteFile(oFile.Path)
		End If
	Next

	' Warn about test mode
	
	If bTest = True Then Call Log("TEST MODE ENABLED. NO FILES WILL BE DELETED!", False)
	
End Sub


' Log(sMessage, bVerbose)
' Write a log entry to the appropriate file. If bVerbose is True then this
' is a verbose log entry and is only written out if bVerboseLog is also
' True. In addition to writing out to a file, it is displayed to the user.

Sub Log(sMessage, bVerbose)

	Dim bNewLog : bNewLog = False

	' Don't log anything if there is no log file defined or this is a
	' verbose log entry and the user has not asked for that.
	
	If sLogName = "" Then Exit Sub
	If bVerbose = True And bVerboseLog <> True Then Exit Sub
	
	Dim sEntry : sEntry = "[" & FormatDateTime(Now(), vbLongTime) & "] " & sMessage

	If fso.FileExists(sLogName) = False Then bNewLog = True
	Dim oLogFile : Set oLogFile = fso.OpenTextFile(sLogName, ForAppending, True)
	
	If bNewLog = True Then
		oLogFile.WriteLine("wmc-dedupe v" & VERSION)
		oLogFile.WriteLine("Copyright © 2011-2014 Richard Lawrence.")
		oLogFile.WriteLine("http://www.fourteenminutes.com/code/wmc-dedupe/")
		oLogFile.WriteLine
		oLogFile.WriteLine("Logfile started on " & Now())
		oLogFile.WriteLine
	End If
	
	' If an empty line is passed to the log file then write just
	' that empty line.
	If sMessage <> "" Then 
		oLogFile.WriteLine(sEntry)
	Else
		oLOgFile.WriteLine
	End If
	
	oLogFile.Close
	Set oLogFile = Nothing

	If sMessage <> "" And bSilent = False Then WScript.Echo(sEntry)
	
End Sub

' Manage_Smaller_File
' Do something with the smaller file.

Sub Manage_Smaller_File(iFile)

	If bDelete = True Then
		Call Log("Deleting " & sFileName(iFile), False)
		If bTest = True Then
			Call Log("TEST MODE: Delete not performed!", False)
			bTestMoved(iFile) = True
		Else
			On Error Resume Next
			fso.DeleteFile sFilePath(iFile), True
			If Err.Number <> 0 Then
				Call Log("Delete failed with error code " & Err.Number & " (" & Err.Description & ")", False)
			End If
		End If
		iDeleted = iDeleted + 1
		Exit Sub
	End If

	' Move the file
	
	If fso.FolderExists(sDuplicateFolder) = False Then
		Call Log("Creating directory: " & sDuplicateFolder, True)
		fso.CreateFolder(sDuplicateFolder)
	End If
	Call Log("Moving " & sFileName(iFile) & " to " & sDuplicateFolder, False)
	If bTest = True Then
		Call Log("TEST MODE: Move not performed!", False)
		bTestMoved(iFile) = True
	Else
		On Error Resume Next
		fso.MoveFile sFilePath(iFile), sDuplicateFolder & "\"
		If Err.Number <> 0 Then
			Call Log("Move failed with error code " & Err.Number & " (" & Err.Description & ")", False)
		End If
		On Error Goto 0
	End If

	iMoved = iMoved + 1
	
End Sub

' Generate_Array
' Load all the data about the files into a set of arrays for much
' faster scanning afterwards.

Sub Generate_Array

	Dim oFolder : Set oFolder = oShell.Namespace(sRecordedTV)
	Dim oFile, iAdded, iSkipped, sTemp, sMessage

	' (GetDetailsOf: 21 = Title, 196 = SubTitle, 259 = Description)
		
	sMessage = "Scanning "
	If bNoWTV = True And bNoDVRMS = True Then sMessage = sMessage & "no "
	If bNoWTV = True And bNoDVRMS = False Then sMessage = sMessage & "DVR-MS "
	If bNoWTV = False And bNoDVRMS = True Then sMessage = sMessage & "WTV "
	If bNoWTV = False And bNoDVRMS = False Then sMessage = sMessage & "WTV and DVR-MS "
	sMessage = sMessage & "files in " & sRecordedTV
	Call Log(sMessage, False)	
	Call Log("Depending on the number of files, this could take a moment...", False)
	
	iAdded = 0 : iSkipped = 0
	For Each oFile in oFolder.Items
		If Is_Recording(oFile.Path) Then
			If bVeryVerboseLog = True Then Call Log("Scanning: " & oFile.Name, True)
			Call Push(sFileName, oFile.Name)
			Call Push(sFilePath, oFile.Path)
			Call Push(iFileSize, fso.GetFile(oFile.Path).Size)  ' Using oFile.Size gives wrong size!?
			Call Push(sFileTitle, Clean_Text(oFolder.GetDetailsOf(oFile, 21), False))
			
			sTemp = Trim(oFolder.GetDetailsOf(oFile, 196))
			If sTemp = "" Then
				' Could be recorded with W8.1MC, in which case it's in a different location
				sTemp = Trim(oFolder.GetDetailsOf(oFile, 204))
			End If
			Call Push(sFileSubTitle, Clean_Text(sTemp, False))
			
			Call Push(bTestMoved, False)
			
			sTemp = Trim(oFolder.GetDetailsOf(oFile, 259))
			If sTemp = "" Then
				' Could be recorded with W8MC, in which case it's in a different location
				sTemp = Trim(oFolder.GetDetailsOf(oFile, 268))
				If sTemp = "" Then
					' Could be recorded with W8.1MC, in which case it's in (yet another) different location
					sTemp = Trim(oFolder.GetDetailsOf(oFile, 272))
				End If
			End If
			
			' Remove some popular strings which start a description. We do this here to make sure
			' that it isn't applied to all strings, only the description.
			sTemp = Remove_First(sTemp, "repeat.")
			sTemp = Remove_First(sTemp, "premier.")
			sTemp = Remove_First(sTemp, "new:")
			sTemp = Remove_First(sTemp, "new.")
			sTemp = Remove_First(sTemp, "new series.")
			sTemp = Remove_First(sTemp, "drama series.")			
			Call Push(sFileDesc, Clean_Text(sTemp, True))
			
			iAdded = iAdded + 1
		Else
			iSkipped = iSkipped + 1
		End If
	Next
	Call Log(iAdded & " media center files found (" & iSkipped & " skipped so " & (iAdded+iSkipped) & " total)", False)
	iTotal = iAdded
	
End Sub

' Push
' Push function for arrays

Sub Push(oArray, oVar) 
 
	Dim iSize
	iSize = GetUBound(oArray) + 1
	Redim Preserve oArray(iSize)
	oArray(iSize) = oVar

End Sub

' GetUBound
' A variant of GetUBound which doesn't fail if the array isn't
' initialised.

Function GetUBound(oArray)

	Dim iSize : iSize = -1

	On Error Resume Next
	iSize = UBound(oArray)
	On Error Goto 0
	GetUBound = iSize

End Function

' Scan_Files
' The main bit of code. Looks at all the files, determines whether any
' of them are duplicates and moves/removes the others.

Sub Scan_Files

	If bMove = False And bDelete = False Then Exit Sub

	Call Generate_Array

	Dim iFileA, iFileB, sMsg, iMax, sMaxUnit
	
	For iFileA = 0 To UBound(sFileName)
		For iFileB = 0 To UBound(sFileName)
			' Make sure both files exist (either physically or virtually as we are using /T)
			If fso.FileExists(sFilePath(iFileA)) And fso.FileExists(sFilePath(iFileB)) And bTestMoved(iFileA) = False And bTestMoved(iFileB) = False Then
				' To ensure that the pair of shows are only checked once
				If sFilePath(iFileA) < sFilePath(iFileB) Then
					' Check if the TV shows are the same
					If Is_Same(iFileA, iFileB) Then
						' They are, so we need to do something with one of them!
						Call Log("Found matching TV shows: " & sFileName(iFileA) & " & " & sFileName(iFileB), False)
		
						iMax = Abs(iFileSize(iFileA)-iFileSize(iFileB))
						SMaxUnit = " bytes"
						If iMax > 1024 Then
							iMax = Int(iMax / 1024)
							sMaxUnit = " KB"
							If iMax > 1024 Then
								iMax = Int(iMax / 1024)
								sMaxUnit = " MB"
								If iMax > 1024 Then
									iMax = Int(iMax / 1024)
									sMaxUnit = " GB"
								End If
							End If
						End If
							
						If iFileSize(iFileA) > iFileSize(iFileB) Then
							Call Log(sFileName(iFileA) & " is larger by " & iMax & sMaxUnit, False)
							Call Manage_Smaller_File(iFileB)
						Else
							Call Log(sFileName(iFileB) & " is larger by " & iMax & sMaxUnit, False)
							Call Manage_Smaller_File(iFileA)
						End If
					End If
				End If
			End If
		Next
	Next
	
	Call Log("Finished processing (" & iMoved & " moved, " & iDeleted & " deleted and " & (iTotal-(iMoved+iDeleted)) & " untouched)", False)
	
End Sub

' Is_Same
' Given two TV shows, work out whether or not they are the same
' programme.

Function Is_Same(iFileA, iFileB)

	Is_Same = False

	' We could use the main array within this function but it gets slightly
	' difficult to understand what is going on and prevents us from easily
	' using a For..Next to cut down on some code.
	
	Dim sTitle(2), sSubTitle(2), sDesc(2), iPos
	sTitle(1) = sFileTitle(iFileA)
	sTitle(2) = sFileTitle(iFileB)
	sSubTitle(1) = sFileSubTitle(iFileA)
	sSubTitle(2) = sFileSubTitle(iFileB)
	sDesc(1) = sFileDesc(iFileA)
	sDesc(2) = sFileDesc(iFileB)
	
	For iPos = 1 to 2
		If iPos = 1 Then
			Call Log("Comparing 1   : " & sFileName(iFileA), True)
		Else
			Call Log("Comparing 2   : " & sFileName(iFileB), True)
		End If
		If bVeryVerboseLog = True Then
			Call Log("Title " & iPos & "       : " & sTitle(iPos), True)
		End If
	Next
	
	For iPos = 1 to 2
		If Len(sTitle(iPos)) = 0 Then
			Call Log("Cannot compare as Title(" & iPos & ") is missing", True)
			Exit Function
		End If
		If Len(sDesc(iPos)) = 0 Then
			Call Log("Cannot compare as Desc(" & iPos & ") is missing", True)
			Exit Function
		End If
		If Len(sSubTitle(iPos)) = 0 Then
			Call Log("Ignoring Sub-title(" & iPos & ") as it is empty", True)
		End If
	Next
			
	' Do the titles match?
	If sTitle(1) <> sTitle(2) Then
		Call Log("Titles do not match", True)
		Exit Function
	End If

	' There is a better chance that these could match, so lets show the
	' details in the log.
	
	If bVeryVerboseLog = True Then
		For iPos = 1 to 2
			Call Log("Sub-title " & iPos & "   : " & sSubTitle(iPos), True)
			Call Log("Description " & iPos & " : " & sDesc(iPos), True)
		Next
	End If

	' Does both of them have subtitles and if so, do they match?
	
	If Len(sSubTitle(1)) > 0 And Len(sSubTitle(2)) > 0 And sSubTitle(1) <> sSubTitle(2) Then
		Call Log("Sub-titles do not match", True)
		Exit Function	
	End If
	
	' Has the subtitle been (stupidly) placed into the description?
	If Len(sSubTitle(2)) = 0 Then
		If sSubTitle(1) & " " & sDesc(1) = sDesc(2) Then
			Is_Same = True
			Call Log("Desc(2) matches Sub-title(1) + Desc(1)", True)
			Exit Function
		End If
	ElseIf Len(sSubtitle(1)) = 0 Then
		If sSubTitle(2) & " " & sDesc(2) = sDesc(1) Then
			Is_Same = True
			Call Log("Desc(1) matches Sub-title(2) + Desc(2)", True)
			Exit Function
		End If
	End If
	
	' Does both of the descriptions match?
	If sDesc(1) <> sDesc(2) Then
		Call Log("Descriptions do not match", True)
		Exit Function
	End If
	
	' Looks like they are the same!
	Is_Same = True
	
End Function

' Clean_Text
' Takes a description of a programme, turns to lower-case, removes all
' punctuation and anything inbetween brackets and square brackets. Also
' removes some common words which could trip up the detection.

Function Clean_Text(sTheText, bRemoveBrackets)

	Dim iPos, iChar, sText

	sText = Clean_Up_Spaces(LCase(sTheText))
	Clean_Text = ""

	' Remove anything inside ( and ) or [ and ].
	
	If bRemoveBrackets = True Then
		Dim iSquare, iBracket : iSquare = 0 : iBracket = 0
		For iPos = 1 to Len(sText)
			iChar = Asc(Mid(sText, iPos, 1))
			Select Case iChar
				Case 40
					iBracket = iBracket + 1
				Case 41
					iBracket = iBracket - 1
				Case 91
					iSquare = iSquare + 1
				Case 93
					iSquare = iSquare - 1
			End Select
			If iBracket = 0 And iSquare = 0 Then
				Clean_Text = Clean_Text & Chr(iChar)
			End If
		Next
		sText = Clean_Up_Spaces(Clean_Text)
		Clean_Text = ""
	End If
	
	' Remove anything non 0-9 or a-z or space
	
	For iPos = 1 to Len(sText)
		iChar = Asc(Mid(sText, iPos, 1))
		If iChar = 32 Or (iChar >= 48 And iChar <= 57) Or (iChar >= 97 And iChar <= 122) Then
			Clean_Text = Clean_Text & Chr(iChar)
		End If
	Next
	
	' Flatten any characters with accents
	Clean_Text = Flatten_Characters(Clean_Text)
	
	Clean_Text = Clean_Up_Spaces(Clean_Text)
	
End Function

' Clean_Up_Spaces
' Simple function to remove leading and ending spaces and keep
' replacing "  " with " " until there are no more to do.

Function Clean_Up_Spaces(sText)

	Clean_Up_Spaces = sText : Exit Function

	Dim sLast : sLast = Trim(sText)
	
	Do
		Clean_Up_Spaces = Trim(Replace(sLast, "  ", " "))
		If Clean_Up_Spaces = sLast Then Exit Function
		sLast = Clean_Up_Spaces
	Loop
	
End Function


' Remove_First(sString, sRemove)
' Given a string (sString), see if it starts with sRemove and, if so,
' remove it.

Function Remove_First(sString, sRemove)

	Dim iLen : iLen = Len(sRemove)
	Dim sClipString : sClipString = Left(sString, 30)
	
	If Len(sClipString) < Len(sString) Then
		sClipString = sClipString & "..."
	End If
	
	If Left(LCase(sString), iLen) = LCase(sRemove) Then
		Remove_First = Trim(Mid(sString, iLen+1, Len(sString)))
		Call Log("Removing '" & sRemove & "' from string '" & sClipString & "'", True)
	Else
		Remove_First = sString
	End If

End Function

' Erase_Duplicates
' Remove any TV shows from the "Duplicates" folder which are over a
' certain number of days.

Sub Erase_Duplicates

	If bErase = False Or iEraseDays <= 0 Then Exit Sub

	Call Log("Deleting duplicate recordings more than " & iEraseDays & " days old in " & sDuplicateFolder, False)
	
	Dim oFiles : Set oFiles = fso.GetFolder(sDuplicateFolder).Files
	Dim oFile, iDeleted, iSkipped : iDeleted = 0 : iSkipped = 0
	For Each oFile In oFiles
	
		If Is_Recording(oFile.Path) = True Then
			' If the recording has not been modified in over a certain number
			' of defined days, then delete it.
		
			Dim iAge : iAge = DateDiff("d", oFile.DateLastModified, Now()) 
			Call Log("Examining duplicate file: " & oFile.Name & " (age: " & iAge & " days)", True)
	
			If iAge > iEraseDays Then
				' Delete this file
				Call Log("Deleting old duplicate file: " & oFile.Name, False)
				iDeleted = iDeleted + 1

				If bTest = True Then
					Call Log("TEST MODE: Delete not performed!", False)
				Else
					On Error Resume Next
					fso.DeleteFile oFile.Path, True
					If Err.Number <> 0 Then
						Call Log("Delete failed with error code " & Err.Number & " (" & Err.Description & ")", False)
					End If
					On Error Goto 0
				End If
			Else
				iSkipped = iSkipped + 1
			End If
		End If
	Next
	
	Call Log("Deleted " & iDeleted & " files and skipped " & iSkipped & " files (" & (iDeleted+iSkipped) & " total)", False)
	
End Sub

' Is_Recording
' Checks whether or not a file is a recorded TV show by looking at the
' extension. Also considers bNoWTV and bNoDVRMS.

Function Is_Recording(sFile)

	Is_Recording = False

	' Don't log it if it turns out to be a directory
	
	If fso.FolderExists(sFile) = True Then
		Call Log("Skipping directory: " & sFile, True)
		Exit Function
	End If

	Dim sExt : sExt = LCase(fso.GetExtensionName(sFile))
	
	' Easy hack to make the rest of the logic easier
	If sExt = "dvr-ms" Then sExt = "dvrms"

	If (sExt = "wtv" And bNoWTV = True) Or (sExt = "dvrms" And bNoDVRMS = True) Then
		Call Log("Skipping " & UCase(sExt) & " file: " & sFile, True)
		Exit Function
	End If
		
	If sExt = "wtv" or sExt = "dvrms" Then
		Is_Recording = True
		Exit Function
	End If		
	
	Call Log("Skipping non-wmc center file: " & sFile, True)
	
End Function

' Check_Not_Recording
' Check whether or not something is being recorded. If it is and the
' script has not been started with /F to force execution, then
' log an error and terminate.

Sub Check_Not_Recording

	Dim bIsRecording : bIsRecording = Running_Process("ehrec.exe")

	If bForceRun = True And bIsRecording = True Then
		Call Log("A recording is in progress. Will continue anyway", False)
		Exit Sub
	ElseIf bForceRun = False And bIsRecording = True Then
		Call Log("A recording is in progress. Stopping", False)
		Call Shutdown
		Exit Sub
	End If
	
	Call Log("No recordings in progress", False)
	
End Sub

' Running_Process(vProcess)
' Determines whether or not a process is running and returns True
' if it is and False if it is not. vProcess can either be a PID
' or the name of the running executable.

Function Running_Process(vProcess)

	Dim objWMIService, objProcess, colProcess, strList

	Running_Process = False
	
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process")

	For Each objProcess in colProcess
		If Trim(LCase(objProcess.Name)) = vProcess Or CStr(objProcess.ProcessID) = vProcess Then
			Call Log("Found running process: " & objProcess.Name & " (pid: " & objProcess.ProcessID & ")", True)
			Running_Process = True
			Exit For
		End If
	Next

	If Running_Process = False Then Call Log("No process matching '" & vProcess & "' found", True)
	
	Set colProcess = Nothing
	Set objWMIService = Nothing
	
End Function

Sub Check_Lock_File

	' If there isn't a lockfile, then we're good

	If fso.FileExists(sLockFile) = False Then
		Call Log("No lock file exists at " & sLockFile, True)
		Call Create_Lock_File
		Exit Sub
	End If

	Call Log("Reading lock file at " & sLockFile, True)
	Dim oFile : Set oFile = fso.OpenTextFile(sLockFile, ForReading)
	Dim sPID : sPID = oFile.ReadLine()
	oFile.Close()
	Set oFile = Nothing
	
	If Trim(sPID) = "" Then
		Call Log("Lock file is empty.", True)
		Call Create_Lock_File
		Exit Sub
	End If

	Call Log("Lock file contains PID " & sPID, True)
	
	If Running_Process(sPID) = False Then
		Call Create_Lock_File
		Exit Sub
	End If
	
	' There is a lock file, it contains a PID and that PID relates
	' to a currently running instance of this script. So we cannot
	' go any further.
	
	Call Log("Another instance of this script is running.", False)
	Call Shutdown
	
End Sub

' Create_Lock_File
' Creates a lock file in %APPDATA%\wmc-dedupe which contains the
' PID of this script. This is used to ensure that we don't run two
' instances of this script at the same time.

Sub Create_Lock_File

	Dim iMyPID : iMyPID = Get_My_Process_ID()
			
	Call Log("Creating lockfile with PID " & iMyPID, True)
	
	Dim oFile : Set oFile = fso.OpenTextFile(sLockFile, ForWriting, True)
	oFile.WriteLine(iMyPID)
	oFile.Close()

End Sub

' Get_My_Process_ID
' Returns the process ID for the currently running script. Many thanks to
' Kul-Tigin at Stack Overflow for the function.

Function Get_My_Process_ID

	Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut
	lOut = 0
	
	Set oShell  = CreateObject("WScript.Shell")
	Set oWMI    = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Randomize

	sCmd = "/K @echo " & Int(Rnd * 3333) * CDbl(Timer) \ 1
	oShell.Run "%comspec% " & sCmd, 0
	WScript.Sleep 100
    
	Set oChldPrcs = oWMI.ExecQuery("Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'", ,32)
	For Each oCols In oChldPrcs
		lOut = oCols.ParentProcessId 'get parent
		oCols.Terminate 'process terminated
		Exit For
	Next
	
	Set oChldPrcs = Nothing
	Set oWMI = Nothing
	Set oShell = Nothing
	
	Get_My_Process_ID = lOut
End Function


' Delete_Lock_File
' Deletes a lock file after sucessful execution

Sub Delete_Lock_File

	If fso.FileExists(sLockFile) = True Then
		Call Log("Deleting lock file at " & sLockFile, True)
		fso.DeleteFile sLockFile, True
	End If
	
End Sub

' Flatten_Characters
' Takes characters such à, é, ï, ô and ü and convert them into non-accented
' alternatives. This means that a show with the word "fiancée" and "fiancee"
' won't be considered different.

Function Flatten_Characters(sText)

	Dim sOut, iCount, iVal
	
	For iCount = 1 To Len(sText)
		iVal = Asc(Mid(sText, iCount, 1))
		
		Select Case iVal
			Case 192,193,194,195,196,197
				sOut = sOut & "A"
			Case 199
				sOut = sOut & "C"
			Case 200,201,202,203
				sOut = sOut & "E"
			Case 204,205,206,207
				sOut = sOut & "I"
			Case 210,211,212,213,214
				sOut = sOut & "O"
			Case 217,218,219,220
				sOut = sOut & "U"
			Case 224,225,226,227,228,229
				sOut = sOut & "a"
			Case 231
				sOut = sOut & "c"
			Case 232,233,234,235
				sOut = sOut & "e"
			Case 236,237,238,239
				sOut = sOut & "i"
			Case 242,243,244,245,246
				sOut = sOut & "o"
			Case 249,250,251,252
				sOut = sOut & "u"
			Case Else
				sOut = sOut & Chr(iVal)
		End Select
	Next
	
	If sOut <> sText Then
		Call Log("Flattened """ & sText & """ to """ & sOut & """", True)
	End If
	
	Flatten_Characters = sOut

End Function