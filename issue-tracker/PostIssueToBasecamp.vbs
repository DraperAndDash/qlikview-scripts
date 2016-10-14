sub PostIssueToBasecamp
	
	Dim strProjectID
	strProjectID = ActiveDocument.Variables("vBasecampProjectID").GetContent().String
	Dim strPostID
	strPostID = ActiveDocument.Variables("vBasecampPostID").GetContent().String
	Dim strBaseURL
	strBaseURL = ActiveDocument.Variables("vBasecampBaseURL").GetContent().String
	
	vURL = strBaseURL + strProjectID + "/posts/" + strPostID + "/comments.xml"
	
	Dim strDate
	strDate = ActiveDocument.Variables("vIssueDate").GetContent().String
	Dim strUser
	strUser = ActiveDocument.Variables("vIssueUser").GetContent().String
	Dim strCompany
	strCompany = ActiveDocument.Variables("vIssueCompany").GetContent().String
	Dim strApp
	strApp = ActiveDocument.Variables("vAppFileName").GetContent().String
	Dim strVersion
	strVersion = ActiveDocument.Variables("vAppVersion").GetContent().String
	Dim strRelease
	strRelease = ActiveDocument.Variables("vAppRelease").GetContent().String
	Dim strSheet
	strSheet = ActiveDocument.Variables("vIssueSheet").GetContent().String
	Dim strObject
	strObject = ActiveDocument.Variables("vIssueObject").GetContent().String
	Dim strTitle
	strTitle = ActiveDocument.Variables("vIssueTitle").GetContent().String
	Dim strDescription
	strDescription = ActiveDocument.Variables("vIssueDescription").GetContent().String

	xmlToSend = "<comment><body><![CDATA["
	xmlToSend = xmlToSend + "<div><b>User:</b> " + strUser + "</div>"
	xmlToSend = xmlToSend + "<div><b>Company:</b> " + strCompany + "</div>"
	xmlToSend = xmlToSend + "<div><b>Date:</b> " + strDate + "</div>"
	xmlToSend = xmlToSend + "<div><b>App Name:</b> " + strApp + " </div>"
	xmlToSend = xmlToSend + "<div><b>Version:</b> " + strVersion + " " + strRelease + " Release </div>"
	xmlToSend = xmlToSend + "<div><b>Sheet:</b> " + strSheet + " </div>"
	xmlToSend = xmlToSend + "<div><b>Title:</b> " + strTitle + " </div>"
	xmlToSend = xmlToSend + "<div><b>Description:</b> " + strDescription + " </div>"
	xmlToSend = xmlToSend + "]]></body></comment>"
	
	set xmlhttp = CreateObject("Microsoft.XMLHTTP")
	xmlhttp.open "POST", vURL, false
	xmlhttp.setRequestHeader "Content-Type", "application/xml"
	xmlhttp.setRequestHeader "User-Agent", "Issue Tracker (www.draperanddash.com/contact)"
	xmlhttp.setRequestHeader "Authorization", "Basic aW5mb0BkcmFwZXJhbmRkYXNoLmNvbTpkMTVjb3ZlcndpdGhEJkQ="
	xmlhttp.send xmlToSend
	
	result = xmlhttp.Status
	
	set issueResponseText = ActiveDocument.Variables("vIssueResponse")
	issueResponseText.setContent result, true

end sub