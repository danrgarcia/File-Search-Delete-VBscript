' VBScript: FileSearch.vbs
' Written by: Daniel Humphries
' Date: 14 August 2015
' Class: COMP230
' Professor: Stanley Kuchel
' ===================================
Set args = WScript.Arguments 'Receive arguments
if WScript.Arguments.Count = 1 then	' Check if there is an argument
	strLocation = WScript.Arguments(0) 'Set drive letter to first argument		
else 'If argument is not received display error message and quit
	WScript.Echo "Error, you must pass argument Scriptname.vbs 'Location'"
	WScript.Quit
end if

' Dim variables
dim count, extFileName, extFile, strFileName
count = 0
extFileName = "Extensions.txt" ' TXT file containing extensions to search for
strFileName = "Outputfile.txt" ' TXT file for writing output

' Read Extension file into str variable
Set FSO = CreateObject("Scripting.FileSystemObject")
Set extFile = FSO.OpenTextFile(extFileName)
strExtensions = extFile.ReadAll
extFile.Close

' Make library file accessible
Set vbsLib = FSO.OpenTextFile("FileAdmin_Lib.vbs",1,False)
librarySubs = vbsLib.ReadAll
vbsLib.Close
Set vbsLib = Nothing
ExecuteGlobal librarySubs

' Message the user what the script is going to do
WScript.Echo "Searching " & strLocation & " for extensions '" & strExtensions & _
"' and writing results to " & strFileName & vbCrlf

' Create object for TXT file output
Set outputFile = FSO.OpenTextFile(strFileName, 2, True)

' Call Method to search for file types
Call FileSearch(strLocation,strExtensions)

' Message user the number of files found
WScript.Echo count & " files found." & vbCrlf
set FSO = Nothing

' Ask if user wants to delete these files
WScript.StdOut.Write("Do you want to delete these files? (Y/N).......................")
answer = WScript.StdIn.Readline()
WScript.StdOut.WriteLine() 
If UCase(answer) = "Y" Then
	Call FileDelete(strFileName)
Else
	WScript.Quit
End If