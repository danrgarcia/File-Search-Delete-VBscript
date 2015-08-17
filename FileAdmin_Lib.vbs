' File administration library

' By value subroutine to search for files through subdirectories
Sub FileSearch(BYVAL strLocation, strExtensions)
    Dim Folder, SubFolder, File
	Dim strExt
	
	Set Folder = FSO.GetFolder(strLocation) ' Create starting folder location object

    For Each File In Folder.Files
        'only proceed if there is an extension on the file.
        If (InStr(File.Name, ".") > 0) Then			
			' Set extensions to uppercase
			For Each strExt in SPLIT(UCASE(strExtensions), ",")
				'If the file's extension is one being searched for, write the path to the output file.
				If Right(UCase(File.Path),LEN(strExt)+1) = "." & strExt Then 
					' Write filepath and filename to output file
					outputFile.WriteLine(file.Path)
					' count how many files have been found
					count = count + 1
					Exit For
				End If
			Next			
		End If
    Next

	' Call the subroutine for each subfolder within the original folder
    For Each SubFolder In Folder.SubFolders
        Call FileSearch(SubFolder.Path, strExtensions)
    Next	
	
End Sub

' Subroutine to delete files listed in file created from search
Sub FileDelete(fileName)
	' Dim variables
	Dim fileToDelete, count
	count = 0
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set fileList = FSO.OpenTextFile(fileName) ' Create object with file passed to the subroutine
	
	' Loop through file and read each line then delete that file
	Do Until fileList.AtEndOfStream 
		fileToDelete = fileList.ReadLine
		FSO.DeleteFile fileToDelete
		count = count + 1 
	Loop
	
	' Message user how many files were deleted
	WScript.Echo count & " files deleted! Files deleted are listed in Outputfile.txt"
	
	' close file
	fileList.Close
End Sub
	