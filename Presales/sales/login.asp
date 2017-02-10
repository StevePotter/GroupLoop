<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->

<p class="Heading" align="center">Salesmen Only</p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if LoggedEmployee() then
	'if we need to send them to a specific page, do it
	if Request("Source") <> "" then Redirect(Request("Source"))

	Session.Timeout = 20
'------------------------End Code-----------------------------
%>
	<p>Hello <%=GetEmployeeFirstName(0)%> (Salesman ID# <b><%=Session("EmployeeID")%></b>). Here are your options:</p>

	<p>
	<b>Refer Friends</b><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="salesman_refer.asp">Make More Money By Referring Friends</a>
	</p>

	<p>
	<b>Sales Material</b><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="tips.asp">Selling Tips</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="material.asp">Material to Distribute</a>
	</p>

	<p>
	<b>Your Information</b><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="salesman_change_info.asp">Change Your Personal Information</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="salesman_statistics.asp">View Your Sales Statistics</a>
	</p>

	<p>
	<b>Support</b><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="salesman_support.asp">Get Help</a>
	</p>
<%
else
	strNickName = Request("NickName")
	strPassword = Request("Password")
	'If they have already tried to log in and failed, print a different message
	if strPassword <> "" or strNickName <> "" then
		Response.Write("<p>Nope, that name and password don't work.  If you aren't a salesman, you must sign up first.  If you are, please try again.</p>")
	else
%>	
		<p>This is the salesmen only section.  If you are a salesman, enter your info to log in.
		<br><a href="info.asp">Become a Salesman</a></p>
<%
	end if
	PrintLogin "login.asp", "Log In"
end if

%>


<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
