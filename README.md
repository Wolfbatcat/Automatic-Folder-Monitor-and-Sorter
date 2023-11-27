# Automatic Folder Monitor & File Sorter

## Fixes and Tweaks
A modified version of xcloudx01's AutoKey script that:
  1. Expands the list of recognized file types to cover more common formats.
  2. Adds categories for spreadsheets, fonts, code, databases, configuration files, and backup files.
  3. Allows for skipping individual files from the script altogether if the word "skip" is detected in the filename (case-insensitive). Useful for mods installed in their archive forms.
  4. Excludes the categorical folders from removal if they are empty.
  5. Moves extracted files to the Processed folder. If DeleteZipFileAfterExtract = 0, the archive will go to the Compressed folder.
  6. Fixed the infinite loop that occurs when DeleteZipFileAfterExtract = 0.
  7. Runs every 10 seconds instead of 60.


------------------

## Overview:

This tool takes the hassle out of manually sorting your downloads folder, or any other folder you tend to just throw files into.
It will move files that fall under a category to a specified subfolder. It will also auto-unzip any zipped files for you.
So you can assign a bulk set of image extensions, and any time the script sees one in your monitored folder, it will move it to your "Images" folder.
Eg: photo.jpg > Photos folder\photo.jpg. ZipFile.7z > ZipFile\ZipFileContents*.*

Fully customizable to suit your needs!


## Initial Setup
1. Download and install 7zip or easy-7Zip (http://www.e7z.org/free-download.htm)
2. Download and install AutoHotKey (https://www.autohotkey.com/download/)
3. Open the .ahk file in a text editor.
4. Change the "MonitoredFolder" value near the top of the script to point to the folder that you want to be monitored for changes.
5. Change "UnzipTo" value to point to where you want zip files to be unzipped to.
5. Change "HowOftenToScanInSeconds" To how often it should check if anything within the folder has changed.
6. Change the "7ZipLocation" to point to where your 7zip's 7z.exe is. Note: If using Easy 7-Zip, the 7z.exe is in the Easy 7-Zip folder, NOT the 7-Zip folder.
7. Under the "Destination folders" section, you can change where files that match a specific file type will be placed. Eg MoveImagesTo = %MonitoredFolder%\Images

## Adding more file types to a category
1. Starting on Line 27, add in the file extension to the list of file extensions on the PushFiletypeToArray command. Eg. "jpeg" and use a comma to seperate the entries.

## Adding custom categories
1. Copy and paste this onto a new line just after the last one on line 32, and adjust the file types within the [ ] brackets, and label what folder they go into at the end to what you'd like to use.

PushFiletypeToArray(FiletypeObjectArray,["exe","msi","cmd"], "FolderNameGoesHere")
