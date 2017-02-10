<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Employees</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=employees_add.asp&ID=" & intID)
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

'We are going to check for errors if they are updating the profile
if Request("Submit") = "Add" then
	strNickName = Format(Request("NickName"))

	if EmployeeNickNameTaken( strNickName ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the nickname " & strNickName & " is already taken."))

	Query = "SELECT * FROM Employees"
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsEmployee.AddNew

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
	<p>The employee has been added. &nbsp;<a href="employees_add.asp">Add another.</a><br>
	<a href="employees_modify.asp">Modify employees.</a>
	</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<form METHOD="post" ACTION="employees_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What is their access level?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "AccessLevel", 0, "Salesman.  Access to salesman section only.  No staff area access.<br>", 1
				PrintRadioOption "AccessLevel", 1, "Basic staff access.  Viewing of customer information, basic maintenance.<br>", 1
				PrintRadioOption "AccessLevel", 2, "Management access.  Includes financial access.<br>", 1
				PrintRadioOption "AccessLevel", 3, "Executive access.<br>", 1
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
				<input type="text" size="40" name="FirstName" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Last Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="LastName" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Nickname
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="NickName" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Password
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Password">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Birthday
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Birthdate", Date, 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Primary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail1" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Secondary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail2">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Beeper (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="Beeper">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="CellPhone">
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
				<input type="text" size="40" name="HomeStreet1" ><br>
				<input type="text" size="40" name="HomeStreet2">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCity">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "HomeState", "" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeZip">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="HomePhone">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCountry">
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
				<input type="text" size="40" name="SecondaryDescription">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryStreet1"><br>
				<input type="text" size="40" name="SecondaryStreet2">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCity">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "SecondaryState", "" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryZip">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="SecondaryPhone">
				&nbsp;&nbsp;&nbsp;ext. <input type="text" size="4" name="SecondaryPExt">				 
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCountry">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
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
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->