<p align="<%=HeadingAlignment%>"><span class=Heading>Enter Your Info To Log In</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>

<%
'-----------------------Begin Code----------------------------
'This is really for the Testdrive an admin/member site.  hard to explain
Session("Admin") = ""

'This logs them in if they need to
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	MemberLogin strPassword, strNickName
end if

if Request.Cookies("SiteNum"&CustomerID)("AutoLogin") = "1" then
	MemberLogin Request.Cookies("SiteNum"&CustomerID)("Password"), Request.Cookies("SiteNum"&CustomerID)("NickName")
end if


strMessage = Request("Message")



strSource = Request("Source")
if strSource = "" then strSource = "index.asp"
strSubmit = Request("Submit")

if not LoggedMember then

	if Request("Password") = "" and Request("NickName") = "" then
		if Request("Type") = "Master" then
			Response.Write "<p>If you are a member, please log in to view the site.  If not, sorry, but this site is strictly for members only.</p>"
			if AllowMemberApplications = 1 then
%>
				<a href="members_apply.asp">Apply for Membership</a>
<%
			end if
		elseif Request("Message") <> "" then
			Response.Write Request("Message")
			if AllowMemberApplications = 1 then
%>
				<br><a href="members_apply.asp">Apply for Membership</a>
<%
			end if
		elseif Request("Action") = "Expired" then
%>
			<p>You need log in to <%=PresVerb%> this <%=Request("Noun")%>.  This is usually due a member leaving the computer for a while, or typing too long.  Your data will saved be after signing in.</p>
<%			if AllowMemberApplications = 1 then
%>
				<a href="members_apply.asp">Apply for Membership</a>
<%
			end if
		else
			Response.Write "<p>This is for members only, so you must log in.</p>"
			if AllowMemberApplications = 1 then
%>
				<a href="members_apply.asp">Apply for Membership</a>
<%
			end if
		end if
	else
		Response.Write "<p>Nope, that's invalid info.  Try again.</p>"
	end if

	if strSubmit = "" then strSubmit = "Log In"

%>
	<p>Please view the GroupLoop.com Terms of Service <a href"http://www.grouploop.com/homegroup/tos.asp">here</a>.  Signing in verifies that you have read and accept the TOS.</p>

	<form METHOD="post" ACTION="<%=SecurePath%>login.asp"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="Source" value="<%=strSource%>">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<input type="hidden" name="ID2" value="<%=Request("ID2")%>">
	<input type="hidden" name="Action" value="<%=Request("Action")%>">
	<input type="hidden" name="Noun" value="<%=Request("Noun")%>">
	<input type="hidden" name="Message" value="<%=Request("Message")%>">
	<input type="hidden" name="PastVerb" value="<%=Request("PastVerb")%>">
	<input type="hidden" name="PresVerb" value="<%=Request("PresVerb")%>">
	<input type="hidden" name="Another" value="<%=Request("Another")%>">

	<% PrintTableHeader 0%> 
		<tr> 
      		<td class="<% PrintTDMain %>" align="right"><%=UsernameLabel%></td>
     		<td class="<% PrintTDMain %>">
				<input type="text" name="NickName" size="30" value="<%=Request("NickName")%>">
			</td>

    	</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Password</td>
      		<td class="<% PrintTDMain %>">
				<input type="password" name="Password" size="30">
			</td>
    	</tr>

		<tr>
      		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="<%=strSubmit %>"><br>
				Remember my password on this computer <input type="checkbox" name="AutoLogin" value="1">
      		</td>
    	</tr>
  	</table>
	</form>

<%
else
	Response.Cookies("SiteNum"&CustomerID).expires = #10/10/2020#
	Response.Cookies("SiteNum"&CustomerID).path = "/"

	if Request("AutoLogin") = "1" then
		Response.Cookies("SiteNum"&CustomerID)("AutoLogin") = "1"
		Response.Cookies("SiteNum"&CustomerID)("NickName") = GetJustNickNameSession()
		Response.Cookies("SiteNum"&CustomerID)("Password") = Session("Password")
	else
		Response.Cookies("SiteNum"&CustomerID)("AutoLogin") = ""
		Response.Cookies("SiteNum"&CustomerID)("NickName") = ""
		Response.Cookies("SiteNum"&CustomerID)("Password") = ""
	end if

	if InStr( strSource, "?") then
		Redirect(strSource & "&ID=" & Request("ID") )
	elseif Request("ID") <> "" then
		Redirect(strSource & "?ID=" & Request("ID") )
	else
		Redirect(strSource)
	end if

end if

'------------------------End Code-----------------------------
%>