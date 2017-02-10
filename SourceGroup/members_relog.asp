<p class="Heading" align="<%=HeadingAlignment%>">Log In As A Different Member</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'
'-----------------------Begin Code----------------------------	
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	MemberLogin strPassword, strNickName
	if LoggedMember then Redirect("members.asp")
end if

strLastName = Request("LastName")
strPassword = Request("Password")
'If they have already tried to log in and failed, print a different message

if strPassword <> "" or strLastName <> "" then
	Response.Write("<p>Nope, that name and password don't work.</p>")
else
%>	<p>Please view the GroupLoop.com Terms of Service <a href"http://www.grouploop.com/homegroup/tos.asp">here</a>.  Signing in verifies that you have read and accept the TOS.</p>
<%
	Session.Abandon
	Response.Write("<p>If you are signed in as a member and want to log in as another, just enter the " & UsernameLabel & " and Password of the member you want to log in as below.</p>")
end if

PrintLogin "members_relog.asp?ID=" & Session("MemberID"), "Log In"
'------------------------End Code-----------------------------
%>
