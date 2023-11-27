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
		MonitoredFolder = C:\Downloads ;Folder to be monitored
		ProcessedFolder = C:\Downloads\Processed ; Extracted files go here
		CompressedFolder = C:\Downloads\Compressed ; Old archives are moved here
		UnzipTo = ProcessedFolder
		HowOftenToScanInSeconds = 10 ;How long we wait before re-scanning the folder for any changes.
		ToolTips = 1 ;Show helper popups showing what the program is doing.
		OverWrite = 1 ;Overwrite duplicate files?
		RemoveEmptyFolders = 1 ;Delete any folders in the monitored folder that are now empty. (recursive)

	;Zip files
		7ZipLocation = C:\Program Files\7-Zip\7z.exe ;Needed to provide unzipping functionality.
		OpenExtractedZip = False ;Open the folder up after extraction has finished?
		DeleteZipFileAfterExtract = 0 ;Recycle the zip file after a successful extract.
		UnzipSuccessSound = 1 ;Play a jingle when unzipped something.

	;What filetypes belong to what group, and what their folder name should be sorted into.
		FiletypeObjectArray := [] ;Array needs to be initiated first to work.
		PushFiletypeToArray(FiletypeObjectArray, ["doc", "docx", "pdf", "txt", "rtf", "odt", "md", "tex", "nfo"], "Documents")
		PushFiletypeToArray(FiletypeObjectArray, ["xls", "xlsx", "ods", "csv"], "Spreadsheets")
		PushFiletypeToArray(FiletypeObjectArray, ["ppt", "pptx", "odp"], "Slides")
		PushFiletypeToArray(FiletypeObjectArray, ["jpg", "jpeg", "png", "gif", "bmp", "tiff", "svg", "ai", "eps", "psd", "tga", "exr"], "Images")
		PushFiletypeToArray(FiletypeObjectArray, ["mp3", "wav", "aac", "flac", "ogg", "aif", "aiff", "mid", "midi"], "Audio")
		PushFiletypeToArray(FiletypeObjectArray, ["mp4", "mov", "wmv", "flv", "avi", "mkv", "mpeg", "mpg", "veg", "prproj", "aep"], "Videos")
		PushFiletypeToArray(FiletypeObjectArray, ["js", "html", "htm", "css", "py", "java", "c", "cpp", "php", "rb", "swift", "json", "pyw", "vbs"], "Code")
		PushFiletypeToArray(FiletypeObjectArray, ["zip", "rar", "tar", "tar.gz", "7z", "r00", "001"], "Compressed")
		PushFiletypeToArray(FiletypeObjectArray, ["exe", "bat", "sh", "msi", "jar", "cmd", "ahk"], "Executable")
		PushFiletypeToArray(FiletypeObjectArray, ["stl", "obj", "fbx", "dae", "3ds", "blend", "max", "ma", "mb"], "3D Models")
		PushFiletypeToArray(FiletypeObjectArray, ["dwg", "dxf", "skp", "step", "stp", "igs", "iges", "ipt", "iam"], "CAD Files")
		PushFiletypeToArray(FiletypeObjectArray, ["svg", "ai", "eps"], "Vector Graphics")
		PushFiletypeToArray(FiletypeObjectArray, ["pdb", "fits"], "Scientific Data")
		PushFiletypeToArray(FiletypeObjectArray, ["iso", "dmg"], "Archiving & Disk Images")
		PushFiletypeToArray(FiletypeObjectArray, ["aep", "prproj", "veg"], "Animation & Video Editing")
		PushFiletypeToArray(FiletypeObjectArray, ["unity", "gmx"], "Game Development")
		PushFiletypeToArray(FiletypeObjectArray, ["step", "stp", "igs", "iges"], "Engineering")
		PushFiletypeToArray(FiletypeObjectArray, ["vmdk", "ova", "ovf"], "Virtual Machines & Containers")

;---------------------------------------------------------------------------------------------------------------------------------------;
; Main
;---------------------------------------------------------------------------------------------------------------------------------------;
;Start the folder monitor
	WaitTimeBetweenScans := HowOftenToScanInSeconds * 1000
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
			global FiletypeObjectArray
			Loop, %Folder%\*, 2, 1
			{
				FL := ((FL<>"") ? "`n" : "" ) A_LoopFileFullPath
				Sort, FL, R D`n ; Arrange folder-paths inside-out
				Loop, Parse, FL, `n
				{
					; Check if the current folder is in the FiletypeObjectArray
					isInArray := false
					for index, obj in FiletypeObjectArray
					{
						if (A_LoopField contains obj.Destination)
						{
							isInArray := true
							break
						}
					}
					; If the current folder is not in the FiletypeObjectArray, remove it
					if (!isInArray)
					{
						FileRemoveDir, %A_LoopField% ; Do not remove the folder unless is  empty
						If ! ErrorLevel
						{
							Del := Del+1,  RFL := ((RFL<>"") ? "`n" : "" ) A_LoopField
							if Tooltips
							{
								Tooltip,Removing empty folder %FL%
								SetTimer, RemoveToolTip, 3000
							}
						}
					}
				}
			}
			return
		}
		
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
			return
		}
		
		UnZip(FileFullName,Dir,FullPath)
		{
			if (InStr(FileFullName, "skip")) ; Skip files with "skip" in the name
				return
			global 7ZipLocation ;Saves having to re-pass this dir each time you use this function.
			global DeleteZipFileAfterExtract
			global OpenExtractedZip
			global Tooltips
			global ProcessedFolder ; Extracted files go here
			global CompressedFolder ; Old archives are moved here
			global UnzipSuccessSound
			
			;Get filename
				StringGetPos,ExtentPos,FileFullName,.,R
				FileName := SubStr(FileFullName,1,ExtentPos)
				if Tooltips
				{
					Tooltip,Unzipping %FileName% > %Dir%\%FileName%
					SetTimer, RemoveToolTip, 3000
				}
				MakeFolderIfNotExist(ProcessedFolder . "\" . FileName)
				Runwait, "%7ZipLocation%" x "%FullPath%" -o"%ProcessedFolder%\%FileName%"
			sleep,2000
			
			IfExist %ProcessedFolder%\%FileName%
			{
				if DeleteZipFileAfterExtract
					Filerecycle, %FullPath%
				else
				{
					MakeFolderIfNotExist(CompressedFolder) ; Ensure the CompressedFolder exists
					FileMove, %FullPath%, %CompressedFolder% ; Move the old archive to the CompressedFolder
				}
				if OpenExtractedZip
					run, %ProcessedFolder%\%FileName%
				if UnzipSuccessSound
					soundplay, *64
			}
			else
				msgbox,,Oh Noes!,Something went wrong and I couldnt unzip %FileName% to %ProcessedFolder%\%FileName%
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
	;Main
		SearchFiles:
		Loop, Files, %MonitoredFolder%\*
		{
			if (InStr(A_LoopFileName, "skip")) ; Skip files with "skip" in the name
				continue
			DestinationFolder := GetDestination(A_LoopFileFullPath)
			if (DestinationFolder = "Compressed")
				UnZip(A_LoopFileName,A_LoopFileDir,A_LoopFileFullPath)
			else if DestinationFolder
			{
				DestinationFolder := MonitoredFolder . "\" . DestinationFolder
				MakeFolderIfNotExist(DestinationFolder)
				FileMove,%A_LoopFileFullPath%,%DestinationFolder%\*.*,%OverWrite% ; *.* is needed else it could be renamed to no extension! (If dest folder failed)
					if Tooltips
					{
						Tooltip,Moving %A_LoopFileName% > %DestinationFolder%
						SetTimer, RemoveToolTip, 3000
					}
			}
		}
		if RemoveEmptyFolders
			RemoveEmptyFolders(MonitoredFolder)
		FindZipFiles(MonitoredFolder,"Processed")
	
	;Other
		RemoveToolTip:
			SetTimer, RemoveToolTip, Off
			ToolTip
		return

^Esc::ExitApp
