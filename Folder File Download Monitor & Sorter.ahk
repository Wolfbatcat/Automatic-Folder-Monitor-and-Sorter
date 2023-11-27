;---------------------------------------------------------------------------------------------------------------------------------------;
; Initialization
;---------------------------------------------------------------------------------------------------------------------------------------;
	#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
	SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
	SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
	#SingleInstance Force
	#Persistent
	#NoTrayIcon

;---------------------------------------------------------------------------------------------------------------------------------------;
; User Variables
;---------------------------------------------------------------------------------------------------------------------------------------;
	;Behaviour
		MonitoredFolder = D:\Downloads
		UnzipTo = D:\Downloads\Compressed
		HowOftenToScanInSeconds = 60 ;How long we wait before re-scanning the folder for any changes.
		ToolTips = 1 ;Show helper popups showing what the program is doing.
		OverWrite = 1 ;Overwrite duplicate files?
		RemoveEmptyFolders = 1 ;Delete any folders in the monitored folder that are now empty. (recursive)

	;Zip files
		7ZipLocation = D:\Program Files\7-Zip\7z.exe ;Needed to provide unzipping functionality.
		OpenExtractedZip = False ;Open the folder up after extraction has finished?
		DeleteZipFileAfterExtract = 1 ;Recycle the zip file after a successful extract.
		UnzipSuccessSound = 1 ;Play a jingle when unzipped something.

	;What filetypes belong to what group, and what their folder name should be sorted into.
		FiletypeObjectArray := [] ;Array needs to be initiated first to work.
		PushFiletypeToArray(FiletypeObjectArray, ["zip","7z","rar","r00","001"], "Compressed")
		PushFiletypeToArray(FiletypeObjectArray, ["jpg", "bmp", "gif", "gifv", "webm", "png", "jpeg", "swf", "tga", "tiff", "exr", "psd"], "Images")
		PushFiletypeToArray(FiletypeObjectArray, ["txt", "nfo", "rtf", "pdf", "docx", "doc", "pptx", "ppt"], "Documents")
		PushFiletypeToArray(FiletypeObjectArray, ["xlsx", "xls", "csv"], "Spreadsheets")
		PushFiletypeToArray(FiletypeObjectArray, ["mp3", "flac", "wav", "ogg"], "Audio")
		PushFiletypeToArray(FiletypeObjectArray, ["avi", "mpg", "mpeg", "mov", "mp4", "mkv", "wmv"], "Videos")
		PushFiletypeToArray(FiletypeObjectArray, ["exe", "msi", "jar", "cmd", "bat", "ahk", "sh"], "Programs")
		PushFiletypeToArray(FiletypeObjectArray, ["ttf", "otf", "woff"], "Fonts")
		PushFiletypeToArray(FiletypeObjectArray, ["cpp", "java", "py", "html", "css", "js"], "Code")
		PushFiletypeToArray(FiletypeObjectArray, ["sqlite", "mdb", "sql"], "Databases")
		PushFiletypeToArray(FiletypeObjectArray, ["ini", "conf", "json", "xml"], "Configuration")
		PushFiletypeToArray(FiletypeObjectArray, ["bak"], "Backup")

;---------------------------------------------------------------------------------------------------------------------------------------;
; Main
;---------------------------------------------------------------------------------------------------------------------------------------;
;Start the folder monitor
	WaitTimeBetweenScans := HowOftenToScanInSeconds * 60
	SetTimer, SearchFiles, %WaitTimeBetweenScans%
	GoSub,SearchFiles ; Immediately do a scan
	return

;---------------------------------------------------------------------------------------------------------------------------------------;
; Functions
;---------------------------------------------------------------------------------------------------------------------------------------;	
	;Utilities
		HasVal(haystack, needle)
		{
			for index, value in haystack
				if (value = needle)
					return index
			if !IsObject(haystack)
				throw Exception("Bad haystack!", -1, haystack)
			return 0
		}
		
		MakeFolderIfNotExist(TheFolderDir)
		{
			ifnotexist,%TheFolderDir%
				FileCreateDir,%TheFolderDir%
		}		
		
		RemoveEmptyFolders(Folder)
		{
		    global Tooltips
		    ExcludedFolders := "Documents Spreadsheets Programs Compressed Videos Images Audio" ; Folders to exclude
		    Loop, %Folder%\*, 2, 1
		    {
		        FL := ((FL<>"") ? "`n" : "" ) A_LoopFileFullPath
		        Sort, FL, R D`n ; Arrange folder-paths inside-out
		        Loop, Parse, FL, `n
		        {
		            FolderToCheck := A_LoopField
		            ; Check if the folder is in the list of excluded folders
		            if (!InStr(ExcludedFolders, FolderToCheck))
		            {
		                FileRemoveDir, %FolderToCheck% ; Do not remove the folder unless it is empty
		                If !ErrorLevel
		                {
		                    Del := Del+1, RFL := ((RFL<>"") ? "`n" : "" ) A_LoopField
		                    if Tooltips
		                    {
		                        Tooltip, Removing empty folder %FL%
		                        SetTimer, RemoveToolTip, 3000
		                    }
		                }
		            }
		        }
		    }
		}
		return
		
		FindZipFiles(Folder,GoalObjectDestination)
		{
			global FiletypeObjectArray
			global MonitoredFolder
			global Tooltips
			i = 0
			
			loop % FiletypeObjectArray.Count() 
			{
				i ++
				;Get a ref to the object that holds the array of extensions we want.
					if FiletypeObjectArray[i].Destination != GoalObjectDestination						
						continue
					o := FiletypeObjectArray[i]
				
				;Unzip
					if o ;Without this it may end up unzipping to the root of C drive? i THINK "" defaults to C:\ when using Loop Files
					{
						Loop, Files, %MonitoredFolder%\*.*,R
						{
							if HasVal(o.Extensions,A_LoopFileExt)
								UnZip(A_LoopFileName,A_LoopFileDir,A_LoopFileFullPath)
						}
					}
			}
		}
		return
		
		UnZip(FileFullName,Dir,FullPath)
		{
			global 7ZipLocation ;Saves having to re-pass this dir each time you use this function.
			global DeleteZipFileAfterExtract
			global OpenExtractedZip
			global Tooltips
			global UnzipTo
			global UnzipSuccessSound
			
			;Get filename
				StringGetPos,ExtentPos,FileFullName,.,R
				FileName := SubStr(FileFullName,1,ExtentPos)
				if Tooltips
				{
					Tooltip,Unzipping %FileName% > %Dir%\%FileName%
					SetTimer, RemoveToolTip, 3000
				}
				MakeFolderIfNotExist(UnzipTo . "\" . FileName)
				Runwait, "%7ZipLocation%" x "%FullPath%" -o"%UnzipTo%\%FileName%"
			sleep,2000
			
			IfExist %UnzipTo%\%FileName%
			{
				if DeleteZipFileAfterExtract
					Filerecycle, %FullPath%
				if OpenExtractedZip
					run, %UnzipTo%\%FileName%
				if UnzipSuccessSound
					soundplay, *64
			}
			else
				msgbox,,Oh Noes!,Something went wrong and I couldn't unzip %FileName% to %UnzipTo%\%FileName%
		}

		
	;Objects
		PushFiletypeToArray(InputArray,FiletypesArray,Destination)
		{
			InputArray.Push(MakeFiletypeObject(FiletypesArray,Destination))
			return InputArray
		}

		MakeFiletypeObject(InputArray,Destination)
		{
			object := []
			object.Extensions := InputArray
			object.Destination := Destination
			return object
		}
		
		GetDestination(TheFile)
		{
			global FiletypeObjectArray
			for i in FiletypeObjectArray
			{
				if HasVal(FiletypeObjectArray[i].Extensions,A_LoopFileExt)
					Destination := FiletypeObjectArray[i].Destination
			}
		return Destination
}

;---------------------------------------------------------------------------------------------------------------------------------------;
; Subroutines
;---------------------------------------------------------------------------------------------------------------------------------------;
	; Main
	    SearchFiles:
	    Loop, Files, %MonitoredFolder%\\*
	    {
	        DestinationFolder := GetDestination(A_LoopFileFullPath)
	        if (DestinationFolder = "Compressed")
	        {
	            ; Check if the file name contains 'skip' (case-insensitive)
	            if InStr(A_LoopFileName, "skip", false, 1)
	                continue  ; Skip this file
	
	            UnZip(A_LoopFileName, A_LoopFileDir, A_LoopFileFullPath)
	        }
	        else if DestinationFolder
	        {
	            DestinationFolder := MonitoredFolder . "\\" . DestinationFolder
	            MakeFolderIfNotExist(DestinationFolder)
	            FileMove,%A_LoopFileFullPath%,%DestinationFolder%\\*.*,%OverWrite%
	            if Tooltips
	            {
	                Tooltip, Moving %A_LoopFileName% > %DestinationFolder%
	                SetTimer, RemoveToolTip, 3000
	            }
	        }
	    }
	    if RemoveEmptyFolders
	        RemoveEmptyFolders(MonitoredFolder)
	    FindZipFiles(MonitoredFolder,"Compressed")
	
	^Esc::ExitApp
