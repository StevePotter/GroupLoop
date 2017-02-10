<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeQuizzes ) or CBool( QuizzesMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_quizzes_add.asp")
if not (LoggedAdmin or CBool( QuizzesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A New Quiz</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------


if Request("Submit") = "Add" then
	if Request("Subject") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddQuiz"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		if Request("Private") = "1" then 
			.Parameters("@IsPrivate") = 1
		else
			.Parameters("@IsPrivate") = 0
		end if

		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Subject") = Format( Request("Subject") )
		.Parameters("@Description") = GetTextArea( Request("Description") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>The quiz has been added. &nbsp;<a href="members_quizques_add.asp?ID=<%=intID%>">Click here</a> to add its questions.</p>
<%
'-----------------------Begin Code----------------------------
else
	Set rsNew = Server.CreateObject("ADODB.Recordset")

	Public DisplayPrivacy

	Query = "SELECT IncludePrivacyQuizzes FROM Look WHERE CustomerID = " & CustomerID
	rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = CBool(rsNew("IncludePrivacyQuizzes")) and not cBool(SiteMembersOnly)

	rsNew.Close
	Set rsNew = Nothing
'------------------------End Code-----------------------------
%>
	<p>After you click 'Add', you will be able to add the questions to this quiz.</p>

	* indicates required information<br>

	<form METHOD="post" ACTION="<%=SecurePath%>members_quizzes_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
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
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Quiz Name</td>
     	<td class="<% PrintTDMain %>"> 
			<input type="text" size="50" name="Subject" >
    	</td>
	</tr>
		<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Quiz Description
		</td>
		<td class="<% PrintTDMain %>">
			<% TextArea "Description", 50, 5, True, "" %>

		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center" colspan="2">
			<input type="submit" name="Submit" value="Add">
		</td>
	</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if
'------------------------End Code-----------------------------
%>
