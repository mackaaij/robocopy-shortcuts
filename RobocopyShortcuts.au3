#Include <File.au3>
#Include <Array.au3>

;Disable tray menu so script cannot be accidentally paused by clicking tray icon
AutoItSetOption("TrayAutoPause",0)

$windowtitle="RobocopyShortcuts 1.1.004"
dim $ErrorTestloop
dim $ErrorInformationTestloop
dim $ErrorRobocopy
dim $ErrorInformationRobocopy

; Set logfile and create it
$logfile=@TempDir & "\RoboCopyShortcuts.log"
_FileCreate($logfile)
_FileWriteLog($logfile,$windowtitle & " log started")

;Command line parameter should point to an .ini file
If $CmdLine[0]=0 then
	MsgBox(4096 + 16,$windowtitle,"Usage: RobocopyShortcuts.exe <path-to-inifile>")
	_FileWriteLog($logfile,"No command line parameter supplied.")
	Exit
EndIf
$inifile=$CmdLine[1]
;TESTING ONLY
;$inifile="RobocopyShortcuts.ini"

; Check if .ini file opened for reading OK
$filehandle = FileOpen($inifile, 0)
If $filehandle = -1 Then
	$varerror = "Unable to open the ini file: " & $inifile
    MsgBox(16,$windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
    Exit
EndIf

;Read the ini file for SubFoldersToCopy
;If this section is not present then NO subfolders will be excluded later on
$SubFoldersToCopy = IniReadSection($inifile,"SubFoldersToCopy")

;Read the GlobalSettings from the ini file
$ShortcutFolder = IniRead($inifile,"GlobalSettings","ShortcutFolder", "")
If $ShortcutFolder = "" Then
	$varerror = "Error occurred reading section [GlobalSettings], variable Source from ini file: " & $inifile
	MsgBox(16, $windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
	Exit
EndIf

;Check for existance of variable TargetFolder and check if it is a local drive
;Users should not use RobocopyShortcuts to copy TO the network
$TargetFolder = IniRead($inifile,"GlobalSettings","TargetFolder", "")
If $TargetFolder = "" Then
	$varerror = "Error occurred reading section [GlobalSettings], variable TargetFolder from ini file: " & $inifile
	MsgBox(16, $windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
	Exit
ElseIf DriveGetType(StringLeft($TargetFolder,3)) <> "Fixed" Then
	$varerror = "Target folder should be a local/fixed drive: " & $TargetFolder & " (now " & DriveGetType(StringLeft($TargetFolder,3)) & ")"
	MsgBox(16, $windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
	Exit
EndIf
; Strip trailing backslashes from $TargetFolder otherwise Robocopy will fail
While StringRight($TargetFolder,1)="\"
	$TargetFolder=StringTrimRight ($TargetFolder,1)
Wend
_FileWriteLog($logfile, "Using TargetFolder (from Section GlobalSettings) as: " & $TargetFolder)

; Check if Robocopy exist (for now checks for itself and does not check inside PATH only working directory)
If NOT FileExists("robocopy.exe") Then
	$varerror = "Robocopy.exe not found. Current working folder: " & @WorkingDir
    MsgBox(16,$windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
	Exit
EndIf

; Read all the .lnk files in the supplied ShortcutFolder into an array
$ShortcutList=_FileListToArray($ShortcutFolder,"*.lnk",1)
If @error Then
	$varerror = "Error occurred reading shortcuts at location: " & $ShortcutFolder
	MsgBox(16, $windowtitle, $varerror)
	_FileWriteLog($logfile, $varerror)
	Exit
EndIf
; Remove first element from array (which contains the number of array items, a counter)
_ArrayDelete($ShortcutList,0)
_FileWriteLog($logfile, "Found the following shortcuts: " & _ArrayToString($ShortcutList,"|"))

ProgressOn($windowtitle, "Processing shortcuts in: " & $ShortcutFolder,"Checking shortcut integrity, please wait...",-1,-1,18)

ProcessShortcuts("test")
If $ErrorTestloop="true" Then
	_FileWriteLog($logfile,$ErrorInformationTestloop)
	MsgBox(4096 + 16,$windowtitle,$ErrorInformationTestloop)
	Exit
EndIf

ProcessShortcuts("copy")

If $ErrorRobocopy="false" Then
	_FileWriteLog($logfile,"Finished without errors.")
	MsgBox(4096 + 64,$windowtitle,"Finished without errors. This box will autoclose in 10 seconds.",10)
Else
	_FileWriteLog($logfile,$ErrorInformationRobocopy)
	MsgBox(4096 + 48,$windowtitle,$ErrorInformationRobocopy)
EndIf

Func ProcessShortcuts($state)	
	If $state="test" Then _FileWriteLog($logfile,"Processing shortcuts in: " & $ShortcutFolder)
		
	; If supplied state = "test" then the supplied fileset will be tested for errors to these can all be fixed by user
	; If supplied state = "copy" then the supplied fileset will be passed to Robocopy
	$ErrorTestloop = "false"
	$ErrorRobocopy = "false"
	$ErrorInformationTestloop = "ERROR - Cannot continue processing shortcuts in: " & $ShortcutFolder
	$ErrorInformationRobocopy = "WARNING - Robocopy returned errors processing shortcuts in: " & $ShortcutFolder & @LF & "(please check free diskspace and access rights)"
	
	;Process each .lnk file and check whether all are folders and accessible
	; Array counter for progressbar
	$i = 0
	For $ShortcutFilename In $ShortcutList
		; Array counter for progressbar
		$i = $i +1
		
		;Read in the path of a shortcut
		$Shortcut = FileGetShortcut($ShortcutFolder & "\" & $ShortcutFilename)
		
		;Formally used to surround path with quotes but then CreateExcludeParameters would not work: $ShortcutPath = """" & $Shortcut[0] & """"
		$ShortcutPath = $Shortcut[0]
		
		; Strip trailing backslashes from $ShortcutPath otherwise Robocopy will fail
		While StringRight($ShortcutPath,1)="\"
			$ShortcutPath=StringTrimRight ($ShortcutPath,1)
		Wend
		
		; Get attributes of $ShortcutPath (used to check for existance and for verifying it's a folder and not a file)
		$attrib = FileGetAttrib($ShortcutPath)
		
		If @error Then 
			$ErrorTestloop="true"
			$ErrorInformationTestloop = $ErrorInformationTestloop & @lf & "Shortcut " & $Shortcut & " supplies a source which cannot be located: " & $ShortcutPath
		Else
			If NOT StringInStr($attrib, "D") Then
				$ErrorTestloop="true"
				$ErrorInformationTestloop = $ErrorInformationTestloop & @lf & "Shortcut " & $Shortcut & " supplies a source which is a file and not a folder: " & $ShortcutPath
			EndIf
		EndIf
		
		$Exclude = CreateExcludeParameter($ShortcutPath)
		
		$TargetFolderName = $TargetFolder & "\" & StringLeft($ShortcutFilename,StringLen($ShortcutFilename)-4)
		
		If $state="copy" And $ErrorTestloop="false" Then
			RoboCopy($i,UBound($ShortcutList),$ShortcutPath,$TargetFolderName,$Exclude)
		EndIf
	Next
EndFunc

Func CreateExcludeParameter($searchfolder)
;This function creates all the exclude parameters for Robocopy
;Unfortunatly Robocopy does not support a parameter for subdirecties to be included but only the reverse
;
;First all subfolders are put into an array
;Then subfolders mentioned in the .ini file are removed from the array
;Finally a string is returned containing all folders to be excluded
;(the .ini file is read at the beginning of the script)
;(no results will return an empty string, resulting in no folders to be excluded)

	;$SubFoldersToCopy should be an array read from the .ini file.
	;If the array does not exist there were no specific directories specified to include so all will be copied (none exluded)
	If NOT IsArray($SubFoldersToCopy) Then Return ""
	
	; Read all the SubFolders into an array
	$SubFolderList=_FileListToArray($searchfolder,"*",2)
	
	; If subfolders are found then delete those that are supposed to be included (read from .ini file) else return ""
	If NOT @Error=1 Then
		; Remove first element from array (which contains the number of array items, a counter)
		_ArrayDelete($SubFolderList,0)
	
		;For each of the SubFoldersToCopy read from the .ini file
		;Skip first array entry because this contains the number of items in the array
		For $i = 1 To $SubFoldersToCopy[0][0]
			;Search for array items and delete them when found
				$Pos = _ArraySearch($SubFolderList,$SubFoldersToCopy[$i][1],0,0,0,0)
				If $Pos <> -1 Then _ArrayDelete($SubFolderList,$Pos)
				;_ArrayDelete returns "" when the array is empty
				If $SubFolderList = "" Then Return ""
		Next
	Else
		Return ""
	EndIf
	
	;Convert the array to a string separated by a pipe (| - not used in filenames)
	$StringArray = _ArrayToString($SubFolderList,"|")
	;If the conversion to a string failed there is only one subfolder
	If @Error = 1 Then
		$StringArray = """" & $SubFolderList[0] & """"
	Else
		;Replace the pipe with " " and surround the whole string with quotes
		$StringArray = """" & StringReplace($StringArray,"|",""" """,0,0) & """"
	EndIf
	;Return the result
	Return($StringArray)
EndFunc

Func RoboCopy($currentfileset,$totalfilesets,$sourcefolder,$destinationfolder,$excludefolders)
 	; Build textstring for progress information (-1 to skip first line which is used for remarks)
	$MainProgressInformation="Processing shortcut " & $currentfileset & " of " & $totalfilesets
	$ProgressInformation = $sourcefolder & " ->" & @LF & $destinationfolder
	
	If NOT FileExists($destinationfolder) Then $ProgressInformation = $ProgressInformation & @lf & "NOTE: Takes a few minutes on the first run!" & @lf
	_FileWriteLog($logfile,$MainProgressInformation)
	ProgressSet(-1,$ProgressInformation,$MainProgressInformation)
	
	$TrayTipInformation=$ProgressInformation & @lf & "Details:"
	If NOT $excludefolders = ""  Then $TrayTipInformation = $TrayTipInformation & @lf & "Folders to exclude: " & $excludefolders
	If $excludefolders = "" Then $TrayTipInformation = $TrayTipInformation & " (none - everything is mirrored)"
		
	;Log detailed information (variable is still called TrayTip because of historical reasons)
	_FileWriteLog($logfile,$TrayTipInformation)
	;TrayTip disabled since it will change every second (annoyance)
	;TrayTip($windowtitle,$TrayTipInformation,-1,17)
	
	;Build parameterlist
	$robocopyparameters=' "' & $sourcefolder & '" "' & $destinationfolder & '"'
	If NOT $excludefolders="" Then $robocopyparameters=$robocopyparameters & ' /XD ' & $excludefolders
	$robocopyparameters=$robocopyparameters & ' /S /PURGE /NP /LOG+:"' & $logfile & '" /R:0 /W:0'
	
	_FileWriteLog($logfile,"Executing: robocopy.exe" & $robocopyparameters)
	
	;Execute Robocopy with parameters
	$val = RunWait("robocopy.exe" & $robocopyparameters, "", @SW_HIDE)
	If $val>7 Then
		$ErrorRobocopy="true"
		$ErrorInformationRobocopy = $ErrorInformationRobocopy & @lf & "Line " & $currentfileset & ", errorlevel " & $val & " (" & $sourcefolder & "->" & $destinationfolder & ")"
	EndIf
	
	;Set copied files to readonly to stress it's for readonly use
	;Robocopy could also do this with "/A+:R" but only for newly copied files
	FileSetAttrib($destinationfolder, "+R", 1)
	
	;Update statistics
	ProgressSet(($currentfileset/$totalfilesets)*100)
	;TrayTip disabled (annoyance) so doesn't have to be reset
	;TrayTip("","",0)
EndFunc