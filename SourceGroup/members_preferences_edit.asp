<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember then Redirect("members.asp?Source=members_preferences_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p class="Heading" align="<%=HeadingAlignment%>">Change Your Preferences</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'We are going to check for errors if they are updating the profile
strSubmit = Request("Submit")
if strSubmit = "Update" or strSubmit = "Update All My Membership Records" or strSubmit = "Update Just This Membership" then

	Query = "SELECT * FROM Members WHERE ID = " & Session("MemberID") & " AND CustomerID = " & CustomerID
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("We can't find your member record."))
	end if

	'We are updating multiple memberships (linked by the CommonID, which must be >0)
	if strSubmit = "Update All My Membership Records" then
		if rsMember("CommonID") = 0 then
			set rsMember = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You don't have any linking membership records."))
		end if
		'Get the linking ID
		intCommonID = rsMember("CommonID")

		'Close then reopen the recordset with multiple recs
		rsMember.Close
		Query = "SELECT * FROM Members WHERE CommonID = " & intCommonID
		rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
		do until rsMember.EOF

			SetData

			rsMember.MoveNext
		loop

	else
		'Set the data, you fucking retard
		SetData
	end if

	'Close dis bitch
	rsMember.Close
	set rsMember = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->
	<p>Your preferences have been changed. &nbsp;<a href="members_preferences_edit.asp">Click here</a> to change them again.</p>
<%
'-----------------------Begin Code----------------------------

else
	Query = "SELECT * FROM Members WHERE ID = " & Session("MemberID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
		if Request.Cookies("SiteNum"&CustomerID)("AutoLogin") = "1" then
			intAutoLog = 1
		else
			intAutoLog = 0
		end if

'------------------------End Code-----------------------------
%>

	<p>Here you can change all your member preferences.</p>

	<form METHOD="post" ACTION="members_preferences_edit.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Automatic Login
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Should you be automatically logged in?<BR><BR>
				
				BUT WAIT!!!  If you are automatically 
				logged in, ANYONE on your computer can access the site.  So, if you are not worried about others using your 
				computer, set this to Yes.  Otherwise, we recommend setting it to No.
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio intAutoLog, "AutoLogin" %>
			</td>
		</tr>
<%
	if IncludeNewsletter then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=NewsletterTitle%> Subscription
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Subscribe to the site newsletter?
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeSiteNewsletter"), "SubscribeSiteNewsletter" %>
			</td>
		</tr>
		
<%
	end if
	if IncludeAnnouncements then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=AnnouncementsTitle%> Subscription
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When an announcement is posted and e-mailed to members, do you want to receive the e-mail?
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeAnnouncements"), "SubscribeAnnouncements" %>
			</td>
		</tr>
		
<%
	end if
	if IncludeMeetings then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=MeetingsTitle%> Subscription
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When a meeting is posted and e-mailed to members, do you want to receive the e-mail?
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeMeetings"), "SubscribeMeetings" %>
			</td>
		</tr>
		
<%
	end if
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				GroupLoop.com Newsletter
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Occasionally GroupLoop will send out a newsletter with company announcements, new features, etc.  Do you want to receive the newsletter?

			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeGroupLoopNewsletter"), "SubscribeGroupLoopNewsletter" %>
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Weekly Site Additions
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want a summary of each week's additions e-mailed to you? 
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeAdditions"), "SubscribeAdditions" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">

<%
	'They have linked records, so allow them to update all their records
	if rsMember("CommonID") > 0 and CommonMember( rsMember("CommonID") ) > 1 then
%>
		<input type="submit" name="Submit" value="Update All My Membership Records"> 
		<input type="submit" name="Submit" value="Update Just This Membership">
<%
	else
%>
		<input type="submit" name="Submit" value="Update">
<%
	end if
%>
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	set rsMember = Nothing
end if



'This member has more than one membership
'The common ID is the ID of their first member record.  Each one after that gets the same commonid
'if the first record is lost, the commonID is still unique, so no big deal
Function CommonMember( intCommonID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumCommonMembers"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CommonID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )

		.Parameters("@CommonID") = intCommonID

		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing
	CommonMember = intCount

End Function


Sub SetData
	Response.Cookies("SiteNum"&CustomerID).expires = #10/10/2020#
	Response.Cookies("SiteNum"&CustomerID).path = "/"
	Response.Cookies("SiteNum"&CustomerID) = Session("MemberID")

	if Request("AutoLogin") = "1" then
		Response.Cookies("SiteNum"&CustomerID)("AutoLogin") = "1"
		Response.Cookies("SiteNum"&CustomerID)("NickName") = GetJustNickNameSession()
		Response.Cookies("SiteNum"&CustomerID)("Password") = Session("Password")
	else
		Response.Cookies("SiteNum"&CustomerID)("AutoLogin") = ""
		Response.Cookies("SiteNum"&CustomerID)("NickName") = ""
		Response.Cookies("SiteNum"&CustomerID)("Password") = ""
	end if


	if Request("SubscribeSiteNewsletter") <> "" then rsMember("SubscribeSiteNewsletter") = Request("SubscribeSiteNewsletter")
	if Request("SubscribeAnnouncements") <> "" then rsMember("SubscribeAnnouncements") = Request("SubscribeAnnouncements")
	if Request("SubscribeMeetings") <> "" then rsMember("SubscribeMeetings") = Request("SubscribeMeetings")
	if Request("SubscribeGroupLoopNewsletter") <> "" then rsMember("SubscribeGroupLoopNewsletter") = Request("SubscribeGroupLoopNewsletter")
	if Request("SubscribeAdditions") <> "" then rsMember("SubscribeAdditions") = Request("SubscribeAdditions")



	rsMember.Update
End Sub
'------------------------End Code-----------------------------
%>
