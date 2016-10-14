sub ClearIssueFields

	set issueTitleField = ActiveDocument.Fields("_IssueTitle")
	issueTitleField.ResetInputFieldValues 0
	
	set issueDescriptionField = ActiveDocument.Fields("_IssueDescription")
	issueDescriptionField.ResetInputFieldValues 0
	
	set issueScreenshotField = ActiveDocument.Fields("_IssueScreenshot")
	issueScreenshotField.ResetInputFieldValues 0

end sub