<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<!-- #include file="..\sourcegroup\media_functions.asp" -->
<!-- #include file="..\sourcegroup\photos_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
Session.Timeout = 20
Server.ScriptTimeout = 5400

Public strError, blProceed, strPath
strError = ""
blProceed = True

Set FileSystem = CreateObject("Scripting.FileSystemObject")
Set upl = Server.CreateObject("SoftArtisans.FileUp")
strPath = GetPath ("posts")
upl.Path = strPath

if upl.Form("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the Section ID."))
intSectionID = CInt(upl.Form("ID"))

if not LoggedMember() and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedMember() then Redirect("members.asp?Source=members_sectionitems_add.asp?ID=" & intSectionID)

Query = "SELECT * FROM Sections WHERE ID = " & intSectionID
Set rsSection = Server.CreateObject("ADODB.Recordset")
rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

if rsSection.EOF then Redirect("error.asp")

Noun = rsSection("NounSingular")
PluralNoun = rsSection("NounPlural")

if rsSection("ModifySecurity") = "Administrators" and not LoggedAdmin() then
	Set rsSection = Nothing
	Redirect("message.asp?Message=" & Server.URLEncode("Sorry" & GetNickNameSession() & ", but only administrators may add " & PluralNoun & "."))
end if

%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add <%=PrintAn(Noun)%>&nbsp;<%=PrintFirstCap(Noun)%></span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span><br>
</p>
<%

'Create the new photo
Set cmdTemp = Server.CreateObject("ADODB.Command")
With cmdTemp
	.ActiveConnection = Connect
	.CommandText = "AddCustomSectionItem"
	.CommandType = adCmdStoredProc

	.Parameters.Refresh

	.Execute , , adExecuteNoRecords
	intItemID = .Parameters("@ID")
End With

Query = "SELECT * FROM CustomSectionItems WHERE ID = " & intItemID
Set rsNew = Server.CreateObject("ADODB.Recordset")
rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect


rsNew("MemberID") = Session("MemberID")
rsNew("ModifiedID") = Session("MemberID")
rsNew("CustomerID") = CustomerID
rsNew("SectionID") = intSectionID
rsNew("IP") = Request.ServerVariables("REMOTE_HOST")


for intFieldNum = 1 to 10
	AddField intFieldNum
next

rsNew.Update
rsNew.Close
Set rsNew = Nothing

rsSection.Close
Set rsSection = Nothing

Set upl = Nothing
Set FileSystem = Nothing



if blProceed and strError = "" then
'------------------------End Code-----------------------------
%>
		<p>The <%=Noun%> has been added. &nbsp;<a href="members_sectionitems_add.asp?ID=<%=intSectionID%>">Click here</a> to add another.<br>
			<a href="section_view.asp?ID=<%=intItemID%>">Click here</a> to view it.
		</p>
<%
'-----------------------Begin Code----------------------------
else
	Query = "DELETE CustomSectionItems WHERE ID " & intItemID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Redirect("incomplete.asp?Message=" & Server.URLEncode(strError))

end if



Set FileSystem = Nothing





Sub AddField( intFieldNum )
	'Make sure this is a valid field

	Response.Write "FieldName"&intFieldNum & "FieldType"&intFieldNum & rsSection("FieldName"&intFieldNum)& "<br>"

	if rsSection("FieldName"&intFieldNum) <> "" and rsSection("FieldType"&intFieldNum) <> "" then
		FieldTitle = rsSection("FieldName"&intFieldNum)
		FieldType = rsSection("FieldType"&intFieldNum)
		FieldName = "Field" & intFieldNum

		if FieldType = "TextSingle" or FieldType = "Link" or FieldType = "Currency" or FieldType = "Option" then
			rsNew(FieldName) = upl.Form(FieldName)
		elseif FieldType = "TextBox" then
			rsNew("FieldLongText"&intFieldNum) = GetTextArea(upl.Form(FieldName))
		elseif FieldType = "Date" then
			rsNew(FieldName) = AssembleDate(FieldName)
		elseif FieldType = "File" then
			SetFile intFieldNum
		elseif FieldType = "Photo" then
			SetPhoto intFieldNum, intItemID
		end if

	end if

End Sub


Sub SetPhoto( intFieldNum, intItemID )
	FieldName = "Field" & intFieldNum

	if upl.Form(FieldName).IsEmpty then
		if rsSection("RequireFieldInput"&i) then
			blProceed = False
			strError = strError & "You did not select a file for " & rsSection("FieldName"&intFieldNum) & "<br>"
		end if
	elseif blProceed then
		'--- Retrieve the file's content type and assign it to a variable
		FTYPE = upl.Form(FieldName).ContentType

		strFileName = ""
		'--- Restrict the file types saved using a Select condition
		if FTYPE = "image/gif" then
			strFileName = "CustomSectionItems"&intItemID&"-"&intFieldNum&".gif"
			upl.Form(FieldName).SaveAs strFileName
			strExt = "gif"
		elseif FTYPE = "image/pjpeg" or FTYPE = "image/jpeg" then
			strFileName = "CustomSectionItems"&intItemID&"-"&intFieldNum&".jpg"
			upl.Form(FieldName).SaveAs strFileName
			strExt = "jpg"
		elseif FTYPE = "image/bmp" then
			strFileName = "CustomSectionItems"&intItemID&"-"&intFieldNum&".bmp"
			upl.Form(FieldName).SaveAs strFileName
			strExt = "bmp"
		else
			upl.Form(FieldName).delete
			blProceed = false
			strError = strError & "You can only upload gif, bitmap (bmp), and jpeg (jpg) images.  Your photo was a banned type of file for " & rsSection("FieldName"&intFieldNum) & ".<br>"
		end if

		'Make the thumbnail
		if strExt <> "gif" and strError = "" then
			strFileName = strPath & strFileName
			strThumbFileName = strPath & "CustomSectionItems"&intItemID&"-"&intFieldNum&"t.jpg"

			'Make sure we have an image to get
			if FileSystem.FileExists (strFileName) then
				CreateThumbnail strFileName, strThumbFileName
				strThumbnailExt = "jpg"
			else
				blProceed = false
				strError = strError & "The photo has been lost and a thumbnail could not be created.  Please try again.<br>"
			end if
		end if
	else
		upl.Form(FieldName).delete
	end if

End Sub




Sub SetFile( intFieldNum )
	FieldName = "Field" & intFieldNum

	'Check on the file
	if upl.Form(FieldName).IsEmpty then
		if rsSection("RequireFieldInput"&i) then
			blProceed = False
			strError = strError & "You did not select a file for " & rsSection("FieldName"&intFieldNum) & "<br>"
		end if

	elseif blProceed then

		'Get rid of the directories and stuff, and get the extension
		strFileName = FormatFileName(Mid(upl.Form(FieldName).UserFilename, InstrRev(upl.Form(FieldName).UserFilename, "\") + 1))
		strExt = GetExtension(strFileName)

		'Make sure it isn't executable
		if lcase(strExt) = ".exe" or lcase(strExt) = ".asp" or lcase(strExt) = ".com" or lcase(strExt) = ".bat" then
			strError = strError & "You are trying to update an invalid type of file."
			blProceed
		end if

		'Double-check for errors
		SetFile intFieldNum

		'We can't have duplicate file names in the folder, so keep adding numbers to the end
		intNum = 1
		do until not FileSystem.FileExists( strPath & strFileName )
			strFileName = GetJustFile( strFileName ) & intNum & "." & GetExtension( strFileName )
			intNum = intNum + 1
		loop

		'Save this badboy file
		upl.Form(FieldName).SaveAs strFileName

		rsNew(FieldName) = strFileName
	else
		upl.Form(FieldName).delete
	end if


End Sub





Function PrintAn(strFollowWord)
	if IsNull( strFollowWord) then
		PrintAn = "a"
	elseif strFollowWord = "" then
		PrintAn = "a"
	else
		testchar = Lcase( Left( strFollowWord,1 ) )

		if Instr( testchar, "aeiou" ) then
			PrintAn = "an"
		else
			PrintAn = "a"
		end if
	end if



End Function


Function PrintFirstCap( strWord )
	PrintFirstCap = UCase( Left(strWord, 1) ) & Right(strWord, Len(strWord)-1)

End Function
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->