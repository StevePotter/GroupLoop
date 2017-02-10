<%	blBypass = true	%>
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
			Response.Write "<p>To use restricted parts of your site, you will need a name and password, which ensures total security.  But since we want you to test drive a site, we have provided you with a name and password.</p>"
	else
		Response.Write "<p>Nope, that's invalid info.  Try again.</p>"
	end if

	if strSubmit = "" then strSubmit = "Log In"

%>
	<p>Just enter "Pastor" in the <%=UsernameLabel%> box and "test" as the password.</p>

	<form METHOD="post" ACTION="<%=SecurePath%>login.asp"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true" ID="Form1">
	<input type="hidden" name="Source" value="<%=strSource%>" ID="Hidden1">
	<input type="hidden" name="ID" value="<%=Request("ID")%>" ID="Hidden2">
	<input type="hidden" name="ID2" value="<%=Request("ID2")%>" ID="Hidden3">
	<input type="hidden" name="Action" value="<%=Request("Action")%>" ID="Hidden4">
	<input type="hidden" name="Noun" value="<%=Request("Noun")%>" ID="Hidden5">
	<input type="hidden" name="Message" value="<%=Request("Message")%>" ID="Hidden6">
	<input type="hidden" name="PastVerb" value="<%=Request("PastVerb")%>" ID="Hidden7">
	<input type="hidden" name="PresVerb" value="<%=Request("PresVerb")%>" ID="Hidden8">
	<input type="hidden" name="Another" value="<%=Request("Another")%>" ID="Hidden9">

	<% PrintTableHeader 0%> 
		<tr> 
      		<td class="<% PrintTDMain %>" align="right"><%=UsernameLabel%></td>
     		<td class="<% PrintTDMain %>">
				<input type="text" name="NickName" size="30" value="Pastor" ID="Text1">
			</td>

    	</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Password</td>
      		<td class="<% PrintTDMain %>">
				<input type="password" name="Password" size="30" ID="Password1" value="test">
			</td>
    	</tr>

		<tr>
      		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="<%=strSubmit %>" ID="Submit1"><br>
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
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->