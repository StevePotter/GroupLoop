<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

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
	EmployeeLogin strPassword, strNickName
end if

strSource = Request("Source")
if strSource = "" then strSource = "index.asp"
strSubmit = Request("Submit")

if not LoggedStaff() then

	if Request("Password") = "" and Request("NickName") = "" then
		Response.Write "<p>Please enter your information to log in.</p>"
	else
		Response.Write "<p>Nope, that's invalid info.  Try again.</p>"
	end if

	if strSubmit = "" then strSubmit = "Log In"

%>
	<form METHOD="post" ACTION="login.asp"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
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
      		<td class="<% PrintTDMain %>" align="right">NickName</td>
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
				<input type="submit" name="Submit" value="<%=strSubmit %>">
      		</td>
    	</tr>
  	</table>
	</form>

<%
else
	if Session("AccessLevel") = 0 then Redirect("message.asp?Message=" & Server.URLEncode("You do not have access to this area."))

	if InStr( strSource, "?") then
		Redirect(strSource & "&ID=" & Request("ID") & "&Submit=" & strSubmit )
	elseif Request("ID") <> "" then
		Redirect(strSource & "?ID=" & Request("ID") & "&Submit=" & strSubmit )
	else
		Redirect(strSource)
	end if

end if

'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->