<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuotes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_quotes_modify.asp")
if not (LoggedAdmin() or CBool( QuotesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Quotes</span><br>
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
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if (blLoggedAdmin and Request("Date") = "") or Request("Author") = "" or Request("Quote") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Author, Quote, Description, Date, IP, ModifiedID FROM Quotes WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Private") = "1" then 
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if
	if blLoggedAdmin then rsUpdate("Date") = Request("Date")
	rsUpdate("Author") = Format( Request("Author") )
	rsUpdate("Quote") = Format( Request("Quote") )
	rsUpdate("Description") = GetTextArea( Request("Description") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The quote has been edited. &nbsp;<a href="members_quotes_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM Quotes WHERE ID = " & intID & " AND " & strMatch
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

	Query = "DELETE Reviews WHERE TargetTable = 'Quotes' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The quote has been deleted. &nbsp;<a href="members_quotes_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Author, Quote, Description, Private FROM Quotes WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_quotes_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* * Person who said the quote</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Author" size="55" value="<%=FormatEdit( rsEdit("Author") )%>">
     	</td>
    </tr>
	<tr> 
   		<td class="<% PrintTDMain %>" align="right" valign="top">* Quote (do not include quotation signs)</td>
    		<td class="<% PrintTDMain %>"> 
   			<textarea name="Quote" cols="55" rows="4" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Quote") )%></textarea>
   		</td>
	</tr>
	<tr> 
   		<td class="<% PrintTDMain %>" align="right" valign="top">Description of quote</td>
   		<td class="<% PrintTDMain %>"> 
			<% TextArea "Description", 55, 4, True, rsEdit("Description") %>
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
	<form METHOD="POST" ACTION="members_quotes_modify.asp">
		View Quotes In The Last  <% PrintDaysOld %>
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
		Query = "SELECT ID, Date, MemberID, Author, Quote, Description FROM Quotes WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Author = rsList("Author")
			Set Description = rsList("Description")
			Set Quote = rsList("Quote")
		intSearchID = SingleSearch()
		Session("SearchID") = intSearchID
		rsList.Close
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="members_quotes_modify.asp">click here</a> to view all quotes.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="members_quotes_modify.asp">click here</a> to view all quotes.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_quotes_modify.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, MemberID, Author, Quote, Private FROM Quotes WHERE CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Author = rsList("Author")
			Set Quote = rsList("Quote")
			Set IsPrivate = rsList("Private")

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
	'They are just cycling through the quotes.  No searching.
	else
		'This is if they requested quotes written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT ID, Date, MemberID, Author, Quote, Private FROM Quotes WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		else
			Query = "SELECT ID, Date, MemberID, Author, Quote, Private FROM Quotes WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set MemberID = rsPage("MemberID")
			Set Author = rsPage("Author")
			Set Quote = rsPage("Quote")
			Set IsPrivate = rsPage("Private")
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_quotes_modify.asp">
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
				<p>Sorry, but there have been no quotes added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there are no quotes at the moment.</p>
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
	GetDesc = UCASE(Author & Quote & ItemDate & GetNickName(MemberID) )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Date</td>
<%		if blLoggedAdmin then %>
		<td class="TDHeader">Author</td>
<%		end if %>
		<td class="TDHeader">Quote</td>
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
	<form METHOD="POST" ACTION="members_quotes_modify.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
<%		if blLoggedAdmin then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
<%		end if %>
		<td class="<% PrintTDMain %>"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("&quot;"&Quote&"&quot; - "&Author)%></a></td>
		<td class="<% PrintTDMain %>"><%=PrintPublic(IsPrivate)%></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this quote, there is no way to get it back.  Are you sure?', 'members_quotes_modify.asp?Submit=Delete&ID=<%=ID%>')">			
			<%if ReviewsExist( "Quotes", ID ) AND blLoggedAdmin then%>
				<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_quotes_modify.asp&TargetTable=Quotes&TargetID=<%=ID%>')">
			<%end if%>	
		</td>
		</tr>
	</form>
<%
End Sub
'------------------------End Code-----------------------------
%>