<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Employees</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=employees_modify.asp&ID=" & Request("ID") & "&Submit=" & Request("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

strSubmit = Request("Submit")

'We are going to check for errors if they are updating the profile
if strSubmit = "Update" then
	strNickName = Format(Request("NickName"))

	if Request("ID") = "" or Request("FirstName") = "" or Request("LastName") = "" or Request("NickName") = "" or Request("Password") = ""	then Redirect("incomplete.asp")

	if UCASE(GetEmployeeNickName(Request("ID"))) <> UCASE(strNickName) and EmployeeNickNameTaken( strNickName ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the nickname " & strNickName & " is already taken."))

	Query = "SELECT * FROM Employees WHERE ID = " & Request("ID")
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsEmployee("AccessLevel") = Request("AccessLevel")

	rsEmployee("FirstName") = Format(Request("FirstName"))
	rsEmployee("LastName") = Format(Request("LastName"))
	rsEmployee("NickName") = strNickName
	rsEmployee("Password") = Request("Password")
	rsEmployee("EMail1") = Request("EMail1")
	rsEmployee("EMail2") = Request("EMail2")
	rsEmployee("Beeper") = Request("Beeper")
	rsEmployee("CellPhone") = Request("CellPhone")
	rsEmployee("Birthdate") = AssembleDate("Birthdate")
	rsEmployee("HomeStreet1") = Request("HomeStreet1")
	rsEmployee("HomeStreet2") = Request("HomeStreet2")
	rsEmployee("HomeCity") = Request("HomeCity")
	rsEmployee("HomeState") = Request("HomeState")
	rsEmployee("HomeZip") = Request("HomeZip")
	rsEmployee("HomeCountry") = Request("HomeCountry")
	rsEmployee("HomePhone") = Request("HomePhone")
	rsEmployee("SecondaryDescription") = Format( Request("SecondaryDescription") )
	rsEmployee("SecondaryStreet1") = Request("SecondaryStreet1")
	rsEmployee("SecondaryStreet2") = Request("SecondaryStreet2")
	rsEmployee("SecondaryCity") = Request("SecondaryCity")
	rsEmployee("SecondaryState") = Request("SecondaryState")
	rsEmployee("SecondaryZip") = Request("SecondaryZip")
	rsEmployee("SecondaryCountry") = Request("SecondaryCountry")
	rsEmployee("SecondaryPhone") = Request("SecondaryPhone")
	rsEmployee("SecondaryPExt") = Request("SecondaryPExt")

	rsEmployee.Update
	rsEmployee.Close
	set rsEmployee = Nothing
'------------------------End Code-----------------------------
%>
	<p>The employee has been edited. &nbsp;<a href="employees_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM Employees WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The employee has been deleted. &nbsp;<a href="employees_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM Employees WHERE ID = " & intID
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsEmployee.EOF then
		set rsEmployee = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<a href="employees_modify.asp?ID=<%=intID%>&Submit=EMail">Click here</a> to e-mail them their name/password.<br>

	<form METHOD="post" ACTION="employees_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=intID%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What is their access level?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "AccessLevel", 0, "Salesman.  Access to salesman section only.  No staff area access.<br>", rsEmployee("AccessLevel")
				PrintRadioOption "AccessLevel", 1, "Basic staff access.  Viewing of customer information, basic maintenance.<br>", rsEmployee("AccessLevel")
				PrintRadioOption "AccessLevel", 2, "Management access.  Includes financial access.<br>", rsEmployee("AccessLevel")
				PrintRadioOption "AccessLevel", 3, "Executive access.<br>", rsEmployee("AccessLevel")
%>
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Name & Such
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* First Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="FirstName" value="<%=rsEmployee("FirstName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Last Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="LastName" value="<%=rsEmployee("LastName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Nickname
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="NickName" value="<%=rsEmployee("NickName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Password
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Password" value="<%=rsEmployee("Password")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Birthday
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Birthdate", rsEmployee("Birthdate"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Primary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail1" value="<%=rsEmployee("EMail1")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Secondary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail2" value="<%=rsEmployee("EMail2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Beeper (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="Beeper" value="<%=rsEmployee("Beeper")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="CellPhone" value="<%=rsEmployee("CellPhone")%>">
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Home Address
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeStreet1" value="<%=rsEmployee("HomeStreet1")%>"><br>
				<input type="text" size="40" name="HomeStreet2" value="<%=rsEmployee("HomeStreet2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCity" value="<%=rsEmployee("HomeCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "HomeState", rsEmployee("HomeState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeZip" value="<%=rsEmployee("HomeZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="HomePhone" value="<%=rsEmployee("HomePhone")%>">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCountry" value="<%=rsEmployee("HomeCountry")%>">
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Secondary Address (optional - school, current residence, etc)
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Description (school, work, beach, etc.)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryDescription" value="<%=rsEmployee("SecondaryDescription")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryStreet1" value="<%=rsEmployee("SecondaryStreet1")%>"><br>
				<input type="text" size="40" name="SecondaryStreet2" value="<%=rsEmployee("SecondaryStreet2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCity" value="<%=rsEmployee("SecondaryCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "SecondaryState", rsEmployee("SecondaryState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryZip" value="<%=rsEmployee("SecondaryZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="SecondaryPhone" value="<%=rsEmployee("SecondaryPhone")%>">
				&nbsp;&nbsp;&nbsp;ext. <input type="text" size="4" name="SecondaryPExt" value="<%=rsEmployee("SecondaryPExt")%>">				 
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCountry" value="<%=rsEmployee("SecondaryCountry")%>">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEmployee.Close
	Set rsEmployee = Nothing
elseif strSubmit = "EMail" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Password, NickName, FirstName, LastName, EMail1 FROM Employees WHERE ID = " & intID
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsEmployee.EOF then
		set rsEmployee = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The member you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strRecipName = rsEmployee("FirstName") & " " & rsEmployee("LastName")
	strNickName = rsEmployee("NickName")
	strEMail = rsEmployee("EMail1")
	strPassword = rsEmployee("Password")

	rsEmployee.Close


	if strEMail = "" then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the Employee does not have an e-mail address entered, so it couldn't be sent."))


	Query = "SELECT Subdirectory, UseDomain, DomainName FROM Customers WHERE ID = " & CustomerID
	rsEmployee.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsEmployee.EOF then
		set rsEmployee = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The customer you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	blUseDomain = CBool(rsEmployee("UseDomain"))
	if blUseDomain then
		strURL = rsEmployee("DomainName")
	else
		strURL = "http://www.GroupLoop.com/" & rsEmployee("Subdirectory")
	end if

	rsEmployee.Close
	set rsEmployee = Nothing

	strSubject = "Your GroupLoop.com site Employeeship info"

	strBody = "Dear " & strRecipName & "," & VbCrLf & VbCrLf & _
	"Here is your Employeeship info for '" & Title & "' in case you lost it: " & VbCrLf & _
	"Your site address: " & strURL & VbCrLf & _
	"Your nickname: " & strNickName & VbCrLf & _
	"Your password: " & strPassword & VbCrLf & VbCrLf & _
	"Please do not respond to this e-mail.  Your questions may be answered by your site administrator(s) " & _
	"or the GroupLoop.com staff." & VbCrLf & _
	"You may e-mail the GroupLoop.com staff at: support@grouploop.com" & VbCrLf & VbCrLf & _
	"                                     Thank you and enjoy," & VbCrLf & _
	"                                         Stephen Potter" & VbCrLf & _
	"                                         President, GroupLoop.com" & VbCrLf & VbCrLf & _
	"Please read GroupLoop.com's Terms Of Service at http://www.GroupLoop.com/homegroup/tos.asp .  Your signing into your site verifies that you have read and accept the Terms Of Service, so please read it carefully." & VbCrLf

	'Set the rest of the mailing info and send it
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = MailerFromName
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.AddRecipient strRecipName, strEMail
	Mailer.Subject    = FormatEdit(strSubject)
	Mailer.BodyText   = strBody

	if not Mailer.SendMail then 
%>		<p>There has been an error, and the email has not been sent to the new Employee.  Please make sure you had a valid e-mail address entered.  Try again, and if the problem
		persists, e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.  Please include the error below.<br>
		Error was '<%=Mailer.Response%>'</p>
<%	else	%>
		<p>The e-mail has been sent. &nbsp;<a href="employees_modify.asp">Click here</a> to modify another Employee.</p>
<%
	end if
	Set Mailer = Nothing
else
	Query = "SELECT ID, FirstName, LastName, NickName FROM Employees ORDER BY LastName"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You have to have employees before you can modify them."))
	end if

	Set ID = rsPage("ID")
	Set FirstName = rsPage("FirstName")
	Set LastName = rsPage("LastName")
	Set NickName = rsPage("NickName")
'-----------------------End Code----------------------------
%>
	<p><a href="employees_add.asp">Add An Employee</a></p>
	<form METHOD="POST" ACTION="employees_modify.asp">
<%
'-----------------------Begin Code----------------------------
	PrintPagesHeader
'-----------------------End Code----------------------------
%>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">Name</td>
		<td class="TDHeader">Nickname</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for p = 1 to rsPage.PageSize
		if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="employees_modify.asp">
		<input type="hidden" name="ID" value="<%=ID%>">
			<tr>
				<td class="<% PrintTDMain %>"><%=FirstName%>&nbsp;<%=LastName%></td>
				<td class="<% PrintTDMain %>"><%=NickName%></td>
				<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this employee, there is no way to get them back.  Are you sure?', 'employees_modify.asp?Submit=Delete&ID=<%=ID%>')">				
				</td>
			</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
		rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
	rsPage.Close
	set rsPage = Nothing
end if

'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->