
'Script for including external files
Sub includeFile(ByVal fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

'Include QVUtils for working with Qlikview Automation API
includeFile "QvUtils.vbs"

'Include SelectFile for selecting the QVW to add the Issue Tracker to
includeFile "SelectFile.vbs"

'Declare all variables
Const ForReading = 1	'Open a file for reading only. You can't write to this file
Const ForWriting = 2	'Open a file for writing
Const ForAppending = 8	'Open a file and write to the end of the file

Dim OpenMessage

Dim strFile
Dim strLine
Dim SheetsStart
Dim SheetsEnd
Dim SheetPositionStart
Dim SheetPositionEnd
Dim SheetNamePositionStart
Dim SheetNamePositionEnd
Dim SheetName
ReDim SheetArray(-1)
Dim AddSheet

Dim SplitStart
Dim SplitEnd
Dim SplitCounter
Dim SplitText
Dim SplitFolder
Dim SplitFileText

Dim AppFullPath
Dim AppFolder
Dim AppName
Dim AppPRJFolder
Dim AppQlikviewProjects
Dim AppQlikviewProjectsFile
Dim AppQlikviewProjectsText
Dim AppModule
Dim AppModuleFile
Dim AppLoadScript
Dim AppLoadScriptFile
Dim AppNumberOfSheets

Dim IssueTrackerXMLObjects
Dim IssueTrackerQlikviewProjects
Dim IssueTrackerQlikviewProjectsFile
Dim IssueTrackerModule
Dim IssueTrackerModuleFile
Dim IssueTrackerLoadScript
Dim IssueTrackerLoadScriptFile



'Instruct user on setting up the application for adding the Issue Tracker Plugin
OpenMessage = "Welcome to the Issue Tracker Plugin installation wizard!" + vbNewLine + vbNewLine
OpenMessage = OpenMessage + "This will add the Issue Tracker Plugin to the requested sheets in your QlikView application enabling users to leave comments to Basecamp." + vbNewLine + vbNewLine
OpenMessage = OpenMessage + "To ensure the plugin can be installed please make sure your QlikView application is in the following state." + vbNewLine + vbNewLine
OpenMessage = OpenMessage + "1. The application is unlocked" + vbNewLine
OpenMessage = OpenMessage + "2. There is no PRJ folder for the application" + vbNewLine
OpenMessage = OpenMessage + "3. You have set the Module security settings to allow all access (Ctrl + M)" + vbNewLine
OpenMessage = OpenMessage + "4. You haven't added the Issue Tracker Plugin to the application previously" + vbNewLine + vbNewLine
OpenMessage = OpenMessage + "NB. Make sure QlikView is open but no Edit scripts or other windows are open within it. If the application haults you may need to go to the QlikView window to help it along." + vbNewLine + vbNewLine
OpenMessage = OpenMessage + "Press OK to select your QlikView application you'd like to add the plugin to."

MsgBox OpenMessage, 0, "Issue Tracker Plugin"

strFile = SelectFile( )

If strFile = "" Then 
    MsgBox("No file selected. Pease select a Qlikview file.")
Else
    AppFullPath = strFile
End If

With New QlikView

	'Set app variables
	' AppFullPath = InputBox("Enter the file path for the application","Issue Tracker Plugin","\\discover\folder_redirection\grfletcher\Desktop\Work\Issue Tracker\Issue Tracker v1_8.qvw")
	AppFolder = Mid(AppFullPath,1,InStrRev(AppFullPath, "\"))
	AppName = Mid(AppFullPath,InStrRev(AppFullPath, "\"), Len(AppFullPath) - InStrRev(AppFullPath, "\") - 3)
	AppPRJFolder = AppFolder + AppName + "-PRJ"
	AppQlikviewProjects = AppPRJFolder + "\QlikviewProject.xml"
	AppLoadScript = AppPRJFolder + "\LoadScript.txt"
	AppModule = AppPRJFolder + "\Module.txt"

	'Define Issue Tracker file paths
	IssueTrackerXMLObjects = "..\Issue Tracker PRJ\Objects XML\*.xml"
	IssueTrackerLoadScript = "..\Issue Tracker PRJ\LoadScript.txt"
	IssueTrackerModule = "..\Issue Tracker PRJ\Module.txt"
	IssueTrackerQlikviewProjects = "..\Issue Tracker PRJ\QlikviewProject.txt"

	'Create File System Object and XMLDOM for Issue Tracker Plugin
	Set IssueTrackerFSO = CreateObject("Scripting.FileSystemObject")

	'Create PRJ folder
	If Not IssueTrackerFSO.FolderExists(AppPRJFolder) Then
		IssueTrackerFSO.CreateFolder AppPRJFolder
	End If

	'Open, Save and Close Qlikview app
	.open(AppFullPath)
	AppNumberOfSheets = .doc.NoOfSheets
	.doc.SaveAs AppFullPath
	.doc.CloseDoc

	'Copy XML Object files over
	IssueTrackerFSO.CopyFile IssueTrackerXMLObjects, AppPRJFolder

	'Append IssueTrackerModule to AppModule
	Set IssueTrackerModuleFile = IssueTrackerFSO.OpenTextFile(IssueTrackerModule)
	Set AppModuleFile = IssueTrackerFSO.OpenTextFile(AppModule, ForAppending, True)

	Do Until IssueTrackerModuleFile.AtEndOfStream
		strLine = IssueTrackerModuleFile.ReadLine()
		AppModuleFile.WriteLine strLine
	Loop

	IssueTrackerModuleFile.Close
	AppModuleFile.Close

	'Append IssueTrackerLoadScript to AppLoadScript
	Set IssueTrackerLoadScriptFile = IssueTrackerFSO.OpenTextFile(IssueTrackerLoadScript)
	Set AppLoadScriptFile = IssueTrackerFSO.OpenTextFile(AppLoadScript, ForAppending, True)

	Do Until IssueTrackerLoadScriptFile.AtEndOfStream
		strLine = IssueTrackerLoadScriptFile.ReadLine()
		AppLoadScriptFile.WriteLine strLine
	Loop

	IssueTrackerLoadScriptFile.Close
	AppLoadScriptFile.Close

	'Read AppQlikviewProjects xml file into a variable
	Set AppQlikviewProjectsFile = IssueTrackerFSO.OpenTextFile(AppQlikviewProjects, ForReading)
	AppQlikviewProjectsText = AppQlikviewProjectsFile.ReadAll
	AppQlikviewProjectsFile.Close

	'Read IssueTracker Qlikview Projects txt file into variable
	Set IssueTrackerQlikviewProjectsFile = IssueTrackerFSO.OpenTextFile(IssueTrackerQlikviewProjects, ForReading)
	IssueTrackerQlikviewProjectsText = IssueTrackerQlikviewProjectsFile.ReadAll
	IssueTrackerQlikviewProjectsFile.Close

	'Get the start and end position of the Sheet xml tag
	SheetsStart = InStr(AppQlikviewProjectsText, "<SHEETS>")+8
	SheetsEnd = InStr(AppQlikviewProjectsText, "</SHEETS>")-1
	SheetPositionStart = SheetsStart
	SheetPositionEnd = SheetsStart

	'Set Split variables
	SplitFolder = "SplitTemp\"
	SplitStart = 1
	SplitCounter = 1

	'Create Split folder for storing temp split files
	IssueTrackerFSO.CreateFolder(SplitFolder)

	'Loop for the number of sheets in the app
	Dim i
	For i = 1 To AppNumberOfSheets

		'Define the start of the Sheet properties, sheet name and Child Ojects. As they are defined in the loop it will move to the next sheet on each loop.
	    SheetPositionStart = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "<PrjSheetProperties>")
		SheetPositionEnd = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "</PrjSheetProperties>")+21
		SheetNamePositionStart = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "<SheetId>Document\")+18
		SheetNamePositionEnd = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "</SheetId>")
		SheetName = Mid(AppQlikviewProjectsText, SheetNamePositionStart, SheetNamePositionEnd - SheetNamePositionStart)
	    ChildObjectsPositionStart = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "<ChildObjects>")+12
		ChildObjectsPositionEnd = InStr(SheetPositionStart+1, AppQlikviewProjectsText, "</ChildObjects>")
		
		'Get confirmation from the user for each sheet if the Issue Tracker should be added
		AddSheet = MsgBox("Add Issue Tracker to " + SheetName +"?",4,"Adding Issue Tracker to each Sheet")

		'If user selects No then tell the user the sheet has not been added
		If AddSheet = 7 Then 
			MsgBox("Issue Tracker has not been added to " + SheetName)
		'If the user says Yes then split the QlikviewProject.xml file at the start of the ChildObjects
		ElseIf AddSheet = 6 Then
			SplitEnd = ChildObjectsPositionStart
			outputFileName = SplitFolder + SheetName + "-" + CStr(SplitCounter) + ".txt"
			SplitText = Mid(AppQlikviewProjectsText, SplitStart, SplitEnd - SplitStart + 2)
			Set outputFile = IssueTrackerFSO.CreateTextFile(outputFileName, True)
			outputFile.Write(SplitText)
			outputFile.Close
			ReDim Preserve SheetArray(UBound(SheetArray) + 1)
			SheetArray(UBound(SheetArray)) = SheetName
			SplitStart = SplitEnd + 2
			SplitCounter = SplitCounter + 1
		End If
	Next

	'Create the End file from the last Sheet split to the end of QlikviewProject.xml file
	outputFileName = "End.txt"
	SplitText = Mid(AppQlikviewProjectsText, SplitStart)
	Set outputFile = IssueTrackerFSO.CreateTextFile(outputFileName, True)
	outputFile.Write(SplitText)
	outputFile.Close

	'Create a new QlikviewProject file and open for appending
	Set NewQlikviewProjectsFile = IssueTrackerFSO.CreateTextFile("NewQlikviewProject.txt")
	NewQlikviewProjectsFile.Close
	Set NewQlikviewProjectsFile = IssueTrackerFSO.OpenTextFile("NewQlikviewProject.txt", ForAppending, True)

	'Add the split files to the new Qlikview Project file and add the IssueTracker objects layouts between each file
	For Each SplitFile in IssueTrackerFSO.GetFolder(SplitFolder).Files
		Set SplitFileObj = IssueTrackerFSO.OpenTextFile(SplitFile.path, ForReading)
		SplitFileText = SplitFileObj.ReadAll
		NewQlikviewProjectsFile.Write(SplitFileText)
		NewQlikviewProjectsFile.Write(IssueTrackerQlikviewProjectsText)
		SplitFileObj.Close
	Next

	'Add the End file to the end off the new QlikviewProject.xml file
	Set IssueTrackerQlikviewProjectsEndFile = IssueTrackerFSO.OpenTextFile("End.txt", ForReading)
	IssueTrackerQlikviewProjectsEndText = IssueTrackerQlikviewProjectsEndFile.ReadAll
	IssueTrackerQlikviewProjectsEndFile.Close

	NewQlikviewProjectsFile.Write(IssueTrackerQlikviewProjectsEndText)
	NewQlikviewProjectsFile.Close

	'Clean up files. Delete split files, delete current AppQlikviewProject file, then move and rename the New QlikviewProject file into the PRJ folder
	IssueTrackerFSO.DeleteFile(SplitFolder+"*")
	' IssueTrackerFSO.DeleteFolder(SplitFolder)
	IssueTrackerFSO.DeleteFile("End.txt")
	IssueTrackerFSO.DeleteFile(AppQlikviewProjects)
	IssueTrackerFSO.MoveFile "NewQlikviewProject.txt", AppQlikviewProjects

	'Open, Reload, Save and Close Qlikview app to bring in the new objects from the PRJ and reload in new variables
	.open(AppFullPath)
	.doc.Reload
	.doc.Save
	.doc.CloseDoc

	'Remove the APp PRJ folder now plugin added
	' IssueTrackerFSO.DeleteFolder(AppPRJFolder)

	'Add Sheets from SheetArray to variable
	Dim strSheets
	For i=1 to UBound(SheetArray)+1
		If Not i = 1 Then 
			strSheets = strSheets + vbNewLine
		End If
		strSheets = strSheets + CStr(i) + ". " + SheetArray(i-1)
	Next

	'Tell the user this has now been completed
	MsgBox(AppName + " has had the Issue Tracker Plugin installed on the following sheets:"+vbNewLine+vbNewLine+strSheets)

	MsgBox("The installation is now complete.")

End With