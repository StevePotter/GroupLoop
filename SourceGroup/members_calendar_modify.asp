<%
'-----------------------Begin Code----------------------------
if not CBool( IncludeCalendar ) then Redirect("error.asp")
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_calendar_modify.asp")
if not (LoggedAdmin or CBool( CalendarMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Events</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" or Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	intID = CInt(Request("ID"))

	StartDate = FormatDateTime( AssembleDate( "Start" ), 2 )
	EndDate = FormatDateTime( AssembleDate( "End" ), 2 )

	if StartDate > EndDate then Redirect("message.asp?Message=" & Server.URLEncode(EndDate & " " & StartDate & "Sorry but your start date comes after the end date.  It just don't make no darn sense."))

	StartTime = TimeValue( CInt(Request("StartHour")) & ":" & CInt(Request("StartMin")) & ":00 " & Request("StartHalf") )
	EndTime = TimeValue( CInt(Request("EndHour")) & ":" & CInt(Request("EndMin")) & ":00 " & Request("EndHalf") )

	if StartDate = EndDate and StartTime > EndTime then Redirect("message.asp?Message=" & Server.URLEncode("Sorry but your event ends before it begins.  Think about it."))

	StartDate = CDate(StartDate & " " & StartTime)
	EndDate = CDate(EndDate & " " & EndTime)


	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM Calendar WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Private") = GetCheckedResult( Request("Private") )

	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate("Date") = AssembleDate( "Date" )
	rsUpdate("StartDate") = StartDate
	rsUpdate("EndDate") = EndDate

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->
	<p>The event has been edited. &nbsp;<a href="members_calendar_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM Calendar WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'Calendar' AND TargetID = " & Request("ID")
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->
	<p>The event has been deleted. &nbsp;<a href="members_calendar_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Public DisplayPrivacy

	Query = "SELECT IncludePrivacyAnnouncements, DisplaySearchAnnouncements, DisplayDaysOldAnnouncements, InfoTextAnnouncements, ListTypeAnnouncements, DisplayDateListAnnouncements, DisplayAuthorListAnnouncements, DisplayPrivacyListAnnouncements  FROM Look WHERE CustomerID = " & CustomerID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = CBool(rsEdit("IncludePrivacyAnnouncements")) and not cBool(SiteMembersOnly)

	rsEdit.Close

	Query = "SELECT ID, Date, Subject, Body, Private, StartDate, EndDate, StartHour, StartMin, EndHour, EndMin FROM Calendar WHERE ID = " & intID & " AND " & strMatch
	rsEdit.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	<a href="inserts_view.asp?Table=InfoPages" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>

	<form method="post" action="<%=SecurePath%>members_calendar_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
<%
		if DisplayPrivacy then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Only let members read it?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Private" value="1" <%=strChecked%>>
			</td>
   		</tr>
<%
		end if
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* Date Created
		</td>
		<td class="<% PrintTDMain %>">
			<% DatePulldown "Date", rsEdit("Date"), 0 %>&nbsp;&nbsp;<% PrintHours "Hour", rsEdit("Date") %><font size="+1">:</font> <% PrintMinutes "Min", rsEdit("Date") %> <% PrintHalf "Half", rsEdit("Date") %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* Start Date
		</td>
		<td class="<% PrintTDMain %>">
			<% DatePulldown "Start", rsEdit("StartDate"), 0 %>&nbsp;&nbsp;<% PrintHours "StartHour", rsEdit("StartDate") %><font size="+1">:</font> <% PrintMinutes "StartMin", rsEdit("StartDate") %> <% PrintHalf "StartHalf", rsEdit("StartDate") %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* End Date
		</td>
		<td class="<% PrintTDMain %>">
			<% DatePulldown "End", rsEdit("EndDate"), 0 %>&nbsp;&nbsp;<% PrintHours "EndHour", rsEdit("EndDate") %><font size="+1">:</font> <% PrintMinutes "EndMin", rsEdit("EndDate") %> <% PrintHalf "EndHalf", rsEdit("EndDate") %>
		</td>
	</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Subject</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Subject" size="55" value="<%=FormatEdit( rsEdit("Subject") )%>">
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Details (inserts allowed)</td>
    	<td class="<% PrintTDMain %>"> 
  				<% TextArea "Body", 55, 15, True, ExtractInserts( rsEdit("Body") ) %>
    	</td>
    </tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

else
'------------------------End Code-----------------------------
%>
	<form METHOD="POST" ACTION="members_calendar_modify.asp">
		View Your Events For  
		<% DatePulldown "Start", Date, 0 %>   
         through 
		<% DatePulldown "End", Date, 0 %>	
		<input type="submit" name="Submit" value="Go">
		<br>
		Or Search For <input type="text" name="Keywords" size="25">
		<input type="submit" name="Submit" value="Go"><br>
	</form>
<%
'-----------------------Begin Code----------------------------
	'Get the searchID from the last page.  May be blank.
	intSearchID = Request("SearchID")

	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" then
		Query = "SELECT ID, Date, MemberID, Subject, Body, StartDate, EndDate, Private FROM Calendar WHERE (" & strMatch & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Body = rsList("Body")
			Set Subject = rsList("Subject")
			Set StartDate = rsList("StartDate")
			Set EndDate = rsList("EndDate")
		intSearchID = SingleSearch()
		Session("SearchID") = intSearchID
		rsList.Close
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="members_calendar_modify.asp">click here</a> to view all your events.</p>
<%
		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_calendar_modify.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, MemberID, Subject, Body, StartDate, EndDate, Private FROM Calendar WHERE " & strMatch & " ORDER BY Date DESC"
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set IsPrivate = rsList("Private")
			Set StartDate = rsList("StartDate")
			Set EndDate = rsList("EndDate")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & rsPage("TargetID")

					PrintTableData

					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing
		end if
	'They are just cycling through the events.  No searching.
	else
		'This is if they requested events written in a time period
		if Request("StartMonth") <> "" then

			StartDate = Request("StartMonth") & "/" & Request("StartDay") & "/" & Request("StartYear")
			EndDate = Request("EndMonth") & "/" & Request("EndDay") & "/" & Request("EndYear")
	
			if not IsDate( StartDate ) or not IsDate( EndDate ) then Redirect("error.asp?Source=members_calendar_modify.asp&Message=Invalid+start+or+end+date.")

			Query = "SELECT ID, Date, MemberID, Subject, Body, StartDate, EndDate, Private FROM Calendar WHERE ( " & strMatch & "AND " & _
					"( (StartDate >='" & StartDate & "' AND StartDate <= '" & EndDate & "') " & _
					"OR " & _
					"(EndDate >='" & StartDate & "' AND EndDate <= '" & EndDate & "') " & _
					"OR " & _
					"(StartDate < '" & StartDate & "' AND EndDate > '" & EndDate & "' ) ) )"  & _
					"ORDER BY StartDate DESC"
		else
			Query = "SELECT ID, Date, MemberID, Subject, Body, StartDate, EndDate, Private FROM Calendar WHERE (" & strMatch & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set MemberID = rsPage("MemberID")
			Set Subject = rsPage("Subject")
			Set IsPrivate = rsPage("Private")
			Set StartDate = rsPage("StartDate")
			Set EndDate = rsPage("EndDate")
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_calendar_modify.asp">
			<input type="hidden" name="StartMonth" value="<%=Request("StartMonth")%>">
			<input type="hidden" name="StartDay" value="<%=Request("StartDay")%>">
			<input type="hidden" name="StartYear" value="<%=Request("StartYear")%>">
			<input type="hidden" name="EndMonth" value="<%=Request("EndMonth")%>">
			<input type="hidden" name="EndDay" value="<%=Request("EndDay")%>">
			<input type="hidden" name="EndYear" value="<%=Request("EndYear")%>">

<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
						PrintTableData
						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
		else
			if Request("StartMonth") <> "" then
'------------------------End Code-----------------------------
%>
				<p>Sorry, but are no events in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>You have to add an event before you can modify it, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		end if
		rsPage.Close
		set rsPage = Nothing
	end if
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(Subject & Body & ItemDate & StartDate & EndDate & GetNickName(MemberID) )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Start Date</td>
		<td class="TDHeader">End Date</td>
<%		if blLoggedAdmin then %>
		<td class="TDHeader">Author</td>
<%		end if %>
		<td class="TDHeader">Subject</td>
		<td class="TDHeader">Public?</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub


'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<form METHOD="POST" ACTION="members_calendar_modify.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(StartDate, 2)%></td>
		<td class="<% PrintTDMain %>" align="center"><%=FormatDateTime(EndDate, 2)%></td>
<%		if blLoggedAdmin then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
<%		end if %>
		<td class="<% PrintTDMain %>"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
		<td class="<% PrintTDMain %>"><%=PrintPublic(IsPrivate)%></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this event, there is no way to get it back.  Are you sure?', 'members_calendar_modify.asp?Submit=Delete&ID=<%=ID%>')">			
			<%if ReviewsExist( "Calendar", ID ) AND blLoggedAdmin then%>
				<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_calendar_modify.asp&TargetTable=Calendar&TargetID=<%=ID%>')">
			<%end if%>	
		</td>
		</tr>
	</form>
<%
End Sub

Sub PrintHours( strName, SelectTime )
	intSelectHour = Hour( SelectTime )

	Response.Write "<select name='" & strName & "' size=1>"

	if intSelectHour = 0 then intSelectHour = 12
	if intSelectHour > 12 then intSelectHour = intSelectHour - 12

	for i = 1 to 12
		strSelected = ""
		if i = intSelectHour then strSelected = " selected"
		Response.Write "<option value='" & i & "'" & strSelected & ">" & i & "</option>"
	next
	Response.Write "</select>"
End Sub

Sub PrintHalf( strName, SelectTime )
	intSelectHour = Hour( SelectTime )

	Response.Write "<select name='" & strName & "' size=1>"
		if intSelectHour < 12 then
			Response.Write "<option value='AM' selected>AM</option>"
			Response.Write "<option value='PM'>PM</option>"
		else
			Response.Write "<option value='AM'>AM</option>"
			Response.Write "<option value='PM' selected>PM</option>"
		end if
	Response.Write "</select>"
End Sub

Sub PrintMinutes( strName, SelectTime)
	intSelectMin = Minute( SelectTime )

	Response.Write "<select name='" & strName & "' size=1>"
	for i = 0 to 59
		strSelected = ""
		if i = intSelectMin then strSelected = " selected"
		stri = ""
		if i < 10 then stri = "0"
		Response.Write "<option value='" & i & "'" & strSelected & ">" & stri & i & "</option>"
	next
	Response.Write "</select>"
End Sub
'------------------------End Code-----------------------------
%>