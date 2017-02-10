<%
'
'-----------------------Begin Code----------------------------
if not ( IncludeNewsletter + NewsletterMembers > 0 ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_newsletter_subscribers.asp")
if not (LoggedAdmin or CBool( NewsletterMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Subscribers</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("EMail") = "" then Redirect("incomplete.asp")

	Query = "SELECT Name, EMail FROM NewsletterSubscribers WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if


	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("EMail") = Format( Request("EMail") )

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The subscriber has been edited. &nbsp;<a href="members_newsletter_subscribers.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM NewsletterSubscribers WHERE ID = " & intID & " AND " & strMatch
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

'------------------------End Code-----------------------------
%>
	<p>The subscriber has been deleted. &nbsp;<a href="members_newsletter_subscribers.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Name, EMail FROM NewsletterSubscribers WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			if (form.EMail.value == ""){
				alert('You forgot the e-mail address.')
				return false;
			}
			return true;
		}

	//-->
	</SCRIPT>
	* indicates required information<br>
	<form method="post" action="members_newsletter_subscribers.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=intID%>">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Name" size="50" value="<%=FormatEdit( rsEdit("Name") )%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* E-mail</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="50" value="<%=FormatEdit( rsEdit("EMail") )%>">
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
	<form METHOD="POST" ACTION="members_newsletter_subscribers.asp">
		View Subscribers In The Last  <% PrintDaysOld %>
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
		Query = "SELECT ID, Name, EMail FROM NewsletterSubscribers WHERE (" & strMatch & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set Name = rsList("Name")
			Set EMail = rsList("EMail")
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
			Try again, or <a href="members_newsletter_subscribers.asp">click here</a> to view all subscribers.</p>
<%
'-----------------------Begin Code----------------------------

		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_newsletter_subscribers.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, EMail, Name FROM NewsletterSubscribers WHERE " & strMatch
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set EMail = rsList("EMail")
			Set Name = rsList("Name")

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
	'They are just cycling through the Newsletters.  No searching.
	else
		'This is if they requested newsletters written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT ID, Date, EMail, Name FROM NewsletterSubscribers WHERE (" & strMatch & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		else
			Query = "SELECT ID, Date, EMail, Name FROM NewsletterSubscribers WHERE (" & strMatch & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set EMail = rsPage("EMail")
			Set Name = rsPage("Name")
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_newsletter_subscribers.asp">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
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
			if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there have been no newsletter subscribers added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there are no newsletter subscribers at the moment.</p>
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
	GetDesc = UCASE(Name & EMail )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Date Subscribed</td>
		<td class="TDHeader">Name</td>
		<td class="TDHeader">EMail</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<form METHOD="POST" ACTION="members_newsletter_subscribers.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<td class="<% PrintTDMain %>"><% Print Name %></td>
		<td class="<% PrintTDMain %>"><%=EMail%></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this subscriber, there is no way to get them back.  Are you sure?', 'members_newsletter_subscribers.asp?Submit=Delete&ID=<%=ID%>')">			
		</td>
		</tr>
	</form>
<%
End Sub
'------------------------End Code-----------------------------
%>