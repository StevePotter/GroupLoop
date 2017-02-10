<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeCalendar ) then Redirect("error.asp")
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_calendar_add.asp")
if not (LoggedAdmin or CBool( CalendarMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add An Event</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
Set rsNew = Server.CreateObject("ADODB.Recordset")


Public DisplayPrivacy

Query = "SELECT IncludePrivacyAnnouncements, DisplaySearchAnnouncements, DisplayDaysOldAnnouncements, InfoTextAnnouncements, ListTypeAnnouncements, DisplayDateListAnnouncements, DisplayAuthorListAnnouncements, DisplayPrivacyListAnnouncements  FROM Look WHERE CustomerID = " & CustomerID
rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
DisplayPrivacy = CBool(rsNew("IncludePrivacyAnnouncements")) and not cBool(SiteMembersOnly)

rsNew.Close

'Add the story
if Request("Submit") = "Add" then
	if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	StartDate = AssembleDate( "Start" )
	EndDate = AssembleDate( "End" )

	if StartDate > EndDate then Redirect("message.asp?Message=" & Server.URLEncode("Sorry but your start date comes after the end date.  It just don't make no darn sense."))

	if Request("StartHour") <> "" and Request("StartMin") <> "" then
		StartTime = TimeValue( CInt(Request("StartHour")) & ":" & CInt(Request("StartMin")) & ":00 " & Request("StartHalf") )
		StartDate = StartDate & " " & StartTime
	else
		StartTime = ""
	end if

	if Request("EndHour") <> "" and Request("EndMin") <> "" then
		EndTime = TimeValue( CInt(Request("EndHour")) & ":" & CInt(Request("EndMin")) & ":00 " & Request("EndHalf") )
		EndDate = EndDate & " " & EndTime
	else
		EndTime = ""
	end if

	if StartDate = EndDate and StartTime > EndTime then Redirect("message.asp?Message=" & Server.URLEncode("Sorry but your event ends before it begins.  Think about it."))


	'If they can add to more than one site....
	if MultiSiteMember() then
		SitesToAdd = Request("SiteCustID")
		'Get the list of sites

		Set rsSites = Server.CreateObject("ADODB.Recordset")

		GetMemberSitesRecordset rsSites

		do until rsSites.EOF
			'If they chose this site to be added to, or all the sites
			if SitesToAdd = "All" or InStr( SitesToAdd, rsSites("CustomerID") ) then
				AddEvent rsSites("CustomerID")
			end if

			rsSites.MoveNext
		loop
		rsSites.Close
		Set rsSites = Nothing
	else
		AddEvent CustomerID
	end if

'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->
	<p>Your event has been added. &nbsp;<a href="members_calendar_add.asp">Click here</a> to add another.<br>
<%
	if intTargetID <> "" then
%>
	<a href="calendar_event_read.asp?ID=<%=intTargetID%>">Click here</a> to view it.
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
	<form method="post" action="<%=SecurePath%>members_calendar_add.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
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
		<tr>
			<td class="<% PrintTDMain %>" align="right">
				Start Date
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Start", Date, 0 %>&nbsp;&nbsp;<% PrintHours "StartHour" %><font size="+1">:</font> <% PrintMinutes "StartMin" %> <% PrintHalf "StartHalf" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" align="right">
				End Date
			</td>

			<td class="<% PrintTDMain %>">
				<% DatePulldown "End", Date, 0 %>&nbsp;&nbsp;<% PrintHours "EndHour" %><font size="+1">:</font> <% PrintMinutes "EndMin" %> <% PrintHalf "EndHalf" %>
			</td>
		</tr>

		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Details (inserts allowed)</td>
    		<td class="<% PrintTDMain %>"> 
 				<% TextArea "Body", 55, 15, True, "" %>
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

Sub PrintHours( strName )
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"


	for i = 1 to 12
		Response.Write "<option value='" & i & "'>" & i & "</option>"
	next
	Response.Write "</select>"
End Sub

Sub PrintHalf( strName )
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"

		Response.Write "<option value='AM'>AM</option>"
		Response.Write "<option value='PM'>PM</option>"
	Response.Write "</select>"
End Sub

Sub PrintMinutes( strName)
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"

	for i = 0 to 59
		stri = ""
		if i < 10 then stri = "0"
		Response.Write "<option value='" & i & "'>" & stri & i & "</option>"
	next
	Response.Write "</select>"
End Sub

Sub AddEvent( intCustID )
	Query = "SELECT ID, Private, MemberID, Subject, Body, CustomerID, IP, ModifiedID, StartDate, EndDate FROM Calendar"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
	rsNew.AddNew
		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Subject") = Format( Request("Subject") )
		rsNew("Body") = GetTextArea( Request("Body") )
		rsNew("Private") = GetCheckedResult( Request("Private") )
		
		rsNew("CustomerID") = intCustID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
		rsNew("StartDate") = StartDate
		rsNew("EndDate") = EndDate
	rsNew.Update
	if intCustID = CustomerID then
		rsNew.MoveNext
		rsNew.MovePrevious
		intTargetID = rsNew("ID")
	end if
	rsNew.Close
End Sub
'------------------------End Code-----------------------------
%>