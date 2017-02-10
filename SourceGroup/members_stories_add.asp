<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeStories ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_Stories_add.asp")
if not (LoggedAdmin() or CBool( StoriesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Story</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'Add the story

blLoggedAdmin = LoggedAdmin()

Set rsNew = Server.CreateObject("ADODB.Recordset")


Public DisplayPrivacy

Query = "SELECT IncludePrivacyStories, DisplaySearchStories, DisplayDaysOldStories, InfoTextStories, ListTypeStories, DisplayDateListStories, DisplayAuthorListStories, DisplayPrivacyListStories  FROM Look WHERE CustomerID = " & CustomerID
rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
DisplayPrivacy = CBool(rsNew("IncludePrivacyStories")) and not cBool(SiteMembersOnly)

rsNew.Close


if Request("Submit") = "Add" then
	if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")

	strSubject = Format( Request("Subject") )
	strBody = GetTextArea( Request("Body") )

	'Get the e-mail shit...
	SendEMail = Request("EMail")
	if SendEMail = "1" then
		SendEMail = true
	else
		SendEMail = false
	end if

	AddNews = Request("AddNews")
	if AddNews = "1" then
		AddNews = true
	else
		AddNews = false
	end if

	if not blLoggedAdmin then AddNews = false


	'Set up the mailer object
	if SendEMail then
		'Set the rest of the mailing info and send it
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.ContentType = "text/html"
		Mailer.IgnoreMalformedAddress = true
		Mailer.RemoteHost  = "mail4.burlee.com"
		Mailer.FromName    = MailerFromName
		Mailer.FromAddress = "support@grouploop.com"
	end if

	'If they can add to more than one site....
	if MultiSiteMember() then
		SitesToAdd = Request("SiteCustID")
		'Get the list of sites

		Set rsSites = Server.CreateObject("ADODB.Recordset")

		GetMemberSitesRecordset rsSites

		do until rsSites.EOF

			SiteCustID = rsSites("CustomerID")
			SiteTitle = rsSites("Title")

			'If they chose this site to be added to, or all the sites
			if SitesToAdd = "All" or InStr( SitesToAdd, SiteCustID ) then
				AddStory SiteCustID
			end if

			rsSites.MoveNext
		loop
		rsSites.Close
		Set rsSites = Nothing
	else
		AddStory CustomerID
	end if 


	if IsObject(Mailer) then Set Mailer = Nothing

	if AddNews then
%>
		<!-- #include file="write_index.asp" -->
<%
	end if

'------------------------End Code-----------------------------
%>
	<p>Your Story has been added. &nbsp;<a href="members_Stories_add.asp">Click here</a> to add another.<br>
<%
	if intTargetID <> "" then
%>
	<a href="Stories_read.asp?ID=<%=intTargetID%>">Click here</a> to read it.
<%
	end if
%>
	</p>
<%
'-----------------------Begin Code----------------------------



else
'------------------------End Code-----------------------------
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Subject.value == "")
				strError += "          You forgot the subject. \n";
			if (form.Body.value == "")
				strError += "          You forgot the details. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can add this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>
	<a href="inserts_view.asp" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>


	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_Stories_add.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
<%
		if DisplayPrivacy then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Only let members read it?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Private" value="1">
			</td>
   		</tr>
<%
		end if
%>
<%
		if MultiSiteMember() then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">What sites should this be added to?  To select more than one, hold down the Control ('Ctrl') key.</td>
			<td class="<% PrintTDMain %>"> 
				<% PrintMemberSites %>
			</td>
   		</tr>
<%
		end if
%>

		<tr> 
			<td class="<% PrintTDMain %>" align="right">Should this Story be sent via e-mail to all the site members?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="EMail" value="1">
			</td>
   		</tr>
<%
		if blLoggedAdmin then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Should this Story be added to the <%=NewsTitle%> as well?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="AddNews" value="1">
			</td>
   		</tr>
<%
		end if
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Details (inserts allowed)</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 20, True, "" %>
    		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if

Set rsNew = Nothing

Sub AddStory( intCustID )
	intCustID = CInt(intCustID)

	Query = "SELECT ID, Private, MemberID, Subject, Body, CustomerID, IP, ModifiedID FROM Stories"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	'Update the fields
	rsNew.AddNew
		if Request("Private") = "1" then 
			rsNew("Private") = 1
		else
			rsNew("Private") = 0
		end if

		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Subject") = strSubject
		rsNew("Body") = strBody
		rsNew("CustomerID") = intCustID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
	rsNew.Update
	if intCustID = CustomerID then
		rsNew.MoveNext
		rsNew.MovePrevious
		intTargetID = rsNew("ID")
	end if
	rsNew.Close


	if AddNews then
		Set cmdTemp = Server.CreateObject("ADODB.Command")
			cmdTemp.ActiveConnection = Connect
			cmdTemp.CommandText = "AddNews"
			cmdTemp.CommandType = adCmdStoredProc
			cmdTemp.Parameters.Refresh

			cmdTemp.Parameters("@ModifiedID") = Session("MemberID")
			cmdTemp.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
			cmdTemp.Parameters("@Body") = strBody
			cmdTemp.Parameters("@CustomerID") = Int(intCustID)
			cmdTemp.Execute , , adExecuteNoRecords
		Set cmdTemp = Nothing
	end if

	if SendEMail then
		'Open up all the members
		rsNew.CacheSize = 50
		Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE ID = " & Session("MemberID")
		rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		Set FirstName = rsNew("FirstName")
		Set LastName = rsNew("LastName")
		Set EMail = rsNew("EMail1")

		'Get the author's info
'		rsNew.Filter = "ID = " & Session("MemberID")

		strAuthor = FirstName & " " & LastName
		strAuthorEMail = EMail

'		rsNew.Filter = ""
		rsNew.Close
		Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE CustomerID = " & intCustID & " AND EMail1 <> '' AND SubscribeStories = 1"
		rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		Set FirstName = rsNew("FirstName")
		Set LastName = rsNew("LastName")
		Set EMail = rsNew("EMail1")


		strTSubject = Title & " - Story by " & strAuthor & " - " & strSubject

		strBody = "This Story was automatically sent to you by " & strAuthor & "'s request.  Please do not respond to this e-mail.  " & _
			"You may reach " & strAuthor & " at <a href='mailto:" & strAuthorEMail & "'>" & strAuthorEMail & "</a><br><br>" & strBody
		'Fix the links to include the grouploop.com so it shows on their mail
		strBody = Replace(strBody, "inserts", "http://www.GroupLoop.com/" & SubDirectory & "/inserts")


		Mailer.Subject = FormatEdit(strTSubject)
		Mailer.BodyText = strBody

		do until rsNew.EOF
			
			Mailer.ClearRecipients
			Mailer.AddRecipient FirstName & " " & LastName, EMail
			Mailer.SendMail

			rsNew.MoveNext
		loop

		rsNew.Close
	end if



End Sub
'------------------------End Code-----------------------------
%>