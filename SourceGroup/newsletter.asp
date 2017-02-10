<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeNewsletter ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>" class=Heading><%=NewsletterTitle%></p>

<%
'-----------------------Begin Code----------------------------
strSubmit = Request("Submit")

if strSubmit = "" then
%>
	What would you like to do?<br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="newsletter.asp?Submit=GoSubscribe">Subscribe to the Newsletter</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="newsletter.asp?Submit=GoUnSubscribe">UnSubscribe to the Newsletter</a><br>
<%
	'Make sure we have newsletters
	if GetNumItems("Newsletters") > 0 then
%>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="newsletter.asp?Submit=Browse">View Past Newsletters</a>
<%
	end if
elseif strSubmit = "GoSubscribe" then

	if LoggedMember and Request("Action") <> "NonMember" then
%>
	<p>To subscribe to the newsletter, you must change your user information.  <a href="members_info_edit.asp">Click here</a> to do it.</p>
	<a href="newsletter.asp?Submit=GoSubscribe&Action=NonMember">Click here</a> to add someone else who isn't a member.
<%
	else
%>
	<p>If you are a member and wish to subscribe, <a href="login.asp?Source=newsletter.asp?Submit=GoSubscribe">click here</a>.</p>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			if (form.EMail.value == ""){
				alert('You forgot your e-mail address.')
				return false;
			}
			return true;
		}

	//-->
	</SCRIPT>
	* indicates required information<br>
	<form method="post" action="newsletter.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Your name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Name" size="50">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Your e-mail address (ex - joespizza@aol.com)</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="50">
     		</td>
		</tr>

		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Subscribe">
    		</td>
		</tr>
  	</table>
	</form>

<%
	end if
elseif strSubmit = "Subscribe" then
	if Request("EMail") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddNewsletterSubscriber"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 500 )
		.Parameters.Append .CreateParameter ("@EMail", adVarWChar, adParamInput, 200 )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@Name") = Format(Request("Name"))
		.Parameters("@EMail") = Format(Request("EMail"))
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing
	if blExists then
%>
	<p>Sorry, but you are already subscribed to our list.</p>
<%
	else
%>
	<p>You are now subscribed to the newsletter.  Thank you!</p>
<%
	end if
elseif strSubmit = "GoUnSubscribe" then

	if LoggedMember then
%>
	<p>To unsubscribe to the newsletter, you must change your user information.  <a href="members_info_edit.asp">Click here</a> to do it.</p>

<%
	else
%>
	<p>If you are a member and wish to unsubscribe, <a href="login.asp?Source=newsletter.asp?Submit=GoUnSubscribe">click here</a>.</p>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			if (form.EMail.value == ""){
				alert('You forgot your e-mail address.')
				return false;
			}
			return true;
		}

	//-->
	</SCRIPT>
	<form method="post" action="newsletter.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	Your e-mail address (ex - joespizza@aol.com) <input type="text" name="EMail" size="30"> <input type="submit" name="Submit" value="UnSubscribe">
	</form>

<%
	end if
elseif strSubmit = "UnSubscribe" then
	if Request("EMail") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp

		.ActiveConnection = Connect
		.CommandText = "DeleteNewsletterSubscriber"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@EMail", adVarWChar, adParamInput, 200 )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@EMail") = Format(Request("EMail"))
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	if blExists then
%>
	<p>You have been removed from the list.  Thank you!</p>
<%
	else
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but your e-mail address wasn't in our records.  Please try again.  AOL users remember the @aol.com!"))
	end if
else
	Query = "SELECT ID, Date, MemberID, Subject FROM Newsletters WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="newsletter.asp">
<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set MemberID = rsPage("MemberID")
		Set Subject = rsPage("Subject")

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
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no past newsletters at the moment.</p>
<%
'-----------------------Begin Code----------------------------
	end if


		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Newsletter&Source=newsletter.asp'>Change Section Options</a></p>"
		end if

	rsPage.Close
	set rsPage = Nothing
end if


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<% if IncludeDate = 1 then %>
		<td class="TDHeader">Date</td>
		<% end if %>	
		<% if IncludeAuthor = 1 then %>
		<td class="TDHeader">Author</td>
		<% end if %>	
		<td class="TDHeader">Subject</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<tr>
		<% if IncludeDate = 1 then %>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<% end if %>	
		<% if IncludeAuthor = 1 then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		<% end if %>	
		<td class="<% PrintTDMainSwitch %>"><a href="newsletter_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
	</tr>
<%
End Sub


'------------------------End Code-----------------------------
%>
