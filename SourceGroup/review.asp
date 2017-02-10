<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
Session.Timeout = 20

strTitle = "Add A Review"
if Request("Title") <> "" then strTitle = Request("Title")
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>"><%=strTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Add the story
if Request("Submit") = "Add" then
	if Request("Body") = "" or Request("Subject") = "" or Request("TargetID") = "" or Request("Table") = "" or Request("Source") = "" or (not LoggedMember AND Request("Author") = "" ) then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddReview"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CustomerID") = CustomerID
		.Parameters("@Author") = ""
		.Parameters("@Email") = ""
		if not LoggedMember then
			.Parameters("@MemberID") = 0
			.Parameters("@Author") = Request("Author")
			.Parameters("@Email") = Request("Email")
			.Parameters("@Subject") = FormatNonMember( Request("Subject") )
			.Parameters("@Body") = GetTextArea( Request("Body") )
		else
			.Parameters("@MemberID") = Session("MemberID")
			.Parameters("@Author") = ""
			.Parameters("@Email") = ""
			.Parameters("@Subject") = Format( Request("Subject") )
			.Parameters("@Body") = GetTextArea( Request("Body") )
		end if
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@TargetTable") = Request("Table")
		.Parameters("@TargetID") = CInt(Request("TargetID"))

		.Execute , , adExecuteNoRecords
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your review has been added. &nbsp;<a href="<%=Request("Source")%>">Click here</a> to view your new review.</p>
<%
'-----------------------Begin Code----------------------------

else
	if Request("Source") = "" or Request("Table") = "" or Request("TargetID") = "" then Redirect("error.asp")
	strSource = Request("Source")
	intID = CInt( Request("TargetID") )
	strTable = Request("Table")

	'If they are nonmembers and clicked so, make sure we know
	if Request("Type") = "NonMember" then Session("NonMember") = "Y"

	'Log in members who typed in their info
	if Request("Password") <> "" or Request("NickName") <> "" then MemberLogin Request("Password"), Request("NickName")

	if Session("NonMember") <> "Y" AND not LoggedMember then
		strLink = "review.asp?Source=" & strSource & "&TargetID=" & intID & "&Table=" & strTable
		if Request("Password") = "" and Request("NickName") = "" then
%>
			<p>If you are a member, please enter your information and log in below. &nbsp;<br><b>If you aren't a member, <a href="<%=strLink%>&Type=NonMember">click here</a> to add your review.</b></p>
<%
		else
%>
			<p>Sorry, but that name and password don't work.  Please try again, or if you aren't a member, <a href="<%=strLink%>&Type=NonMember">click here</a> to add your review.</b></p>
<%
		end if
		PrintLogin strLink, "Log In"
	else

'------------------------End Code-----------------------------
%>
		* indicates required information<br>
		<form method="post" action="review.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
<%		if LoggedMember then	%>
			<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
			<input type="hidden" name="Password" value="<%=Session("Password")%>">
<%		end if	%>
		<input type="hidden" name="Source" value="<%=strSource%>">
		<input type="hidden" name="Table" value="<%=strTable%>">
		<input type="hidden" name="TargetID" value="<%=intID%>">
		<% PrintTableHeader 0 %>
<%		if not LoggedMember then	%>
				<tr>
					<td class="<% PrintTDMain %>" align="right">
						* Your Name
					</td>
					<td class="<% PrintTDMain %>">
						<input type="text" size="25" name="Author">
					</td>
				</tr>
				<tr>
					<td class="<% PrintTDMain %>" align="right">
						Your E-Mail
					</td>
					<td class="<% PrintTDMain %>">
						<input type="text" size="25" name="EMail">
					</td>
				</tr>
<%		end if %>
			<tr>
				<td class="<% PrintTDMain %>" align="right">
					* Headline for the review
				</td>
				<td class="<% PrintTDMain %>">
					<input type="text" size="25" name="Subject">
				</td>
			</tr>
			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="top">* Review</td>
    			<td class="<% PrintTDMain %>"> 
					<% TextArea "Body", 55, 8, True, "" %>
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
end if
'------------------------End Code-----------------------------
%>