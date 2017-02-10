<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=salesman_change_info.asp")
%>

<p align="center"><span class=Heading>Change Your Personal Information</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<%
'-----------------------Begin Code----------------------------
'We are going to check for errors if they are updating the profile
strSubmit = Request("Submit")
if strSubmit = "Update" then
	strNickName = Format(Request("NickName"))

	if Request("FirstName") = "" or Request("LastName") = "" or strNickName = "" or Request("Password") = "" then Redirect("incomplete.asp")

	if UCase(GetEmployeeNickName(Session("EmployeeID"))) <> UCase(strNickName) AND EmployeeNickNameTaken( strNickName ) then
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but that nickname is already taken by another salesman.  Try another nickname."))
	end if

	Query = "SELECT * FROM Employees WHERE ID = " & Session("EmployeeID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("We can't find your salesman record."))
	end if

	SetData

	'Set the new PW
	Session("Password") = Request("Password")

	'Close dis bitch
	rsMember.Close
	set rsMember = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your info has been edited. &nbsp;<a href="salesman_change_info.asp">Click here</a> to edit it again.</p>
<%
'-----------------------Begin Code----------------------------

else
	Query = "SELECT * FROM Employees WHERE ID = " & Session("EmployeeID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
'------------------------End Code-----------------------------
%>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.FirstName.value == "")
				strError += "          You forgot your first name. \n";
			if (form.LastName.value == "")
				strError += "          You forgot your last name. \n";
			if (form.NickName.value == "")
				strError += "          You forgot your nickname. \n";
			if (form.Password.value == "")
				strError += "          You forgot your password. \n";
			if (form.EMail1.value == "")
				strError += "          You forgot your e-mail address. \n";
			if (form.HomeStreet1.value == "")
				strError += "          You forgot your home street. \n";	
			if (form.HomeCity.value == "")
				strError += "          You forgot your home city. \n";				
			if (form.HomeState.value == "")
				strError += "          You forgot your home state. \n";				
			if (form.HomeZip.value == "")
				strError += "          You forgot your home zip code. \n";				
			if (form.HomeCountry.value == "")
				strError += "          You forgot your home country. \n";				
			if (form.HomePhone.value == "")
				strError += "          You forgot your home phone number. \n";				
				
			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can add this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>

	<p>Here you can change all your information  Please try to keep everything current.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>salesman_change_info.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
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
				<input type="text" size="40" name="FirstName" value="<%=rsMember("FirstName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Last Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="LastName" value="<%=rsMember("LastName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Nickname
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="NickName" value="<%=rsMember("NickName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Password
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Password" value="<%=rsMember("Password")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Birthday
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Birthdate", rsMember("Birthdate"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Primary E-Mail Address
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail1" value="<%=rsMember("EMail1")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Secondary E-Mail Address
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail2" value="<%=rsMember("EMail2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Page (http://www.yourpage.com)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="WebSite" value="<%=rsMember("WebSite")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Beeper (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="Beeper" value="<%=rsMember("Beeper")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="CellPhone" value="<%=rsMember("CellPhone")%>">
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Home Address
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeStreet1" value="<%=rsMember("HomeStreet1")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Street (2nd line)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeStreet2" value="<%=rsMember("HomeStreet2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCity" value="<%=rsMember("HomeCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "HomeState", rsMember("HomeState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeZip" value="<%=rsMember("HomeZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="HomePhone" value="<%=rsMember("HomePhone")%>">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCountry" value="<%=rsMember("HomeCountry")%>">
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
				<input type="text" size="40" name="SecondaryDescription" value="<%=rsMember("SecondaryDescription")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryStreet1" value="<%=rsMember("SecondaryStreet1")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street (2nd line)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryStreet2" value="<%=rsMember("SecondaryStreet2")%>">
			</td>
		</tr>
				<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCity" value="<%=rsMember("SecondaryCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "SecondaryState", rsMember("SecondaryState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryZip" value="<%=rsMember("SecondaryZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="SecondaryPhone" value="<%=rsMember("SecondaryPhone")%>">
				&nbsp;&nbsp;&nbsp;ext. <input type="text" size="4" name="SecondaryPExt" value="<%=rsMember("SecondaryPExt")%>">				 
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCountry" value="<%=rsMember("SecondaryCountry")%>">
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
	rsMember.Close
	Set rsMember = Nothing
end if




Sub SetData
	rsMember("FirstName") = Format(Request("FirstName"))
	rsMember("LastName") = Format(Request("LastName"))
	rsMember("NickName") = strNickName
	rsMember("Password") = Request("Password")
	rsMember("EMail1") = Request("EMail1")
	rsMember("EMail2") = Request("EMail2")
	rsMember("WebSite") = Request("WebSite")
	rsMember("Beeper") = Format(Request("Beeper"))
	rsMember("CellPhone") = Format(Request("CellPhone"))
	rsMember("Birthdate") = AssembleDate("Birthdate")
	rsMember("HomeStreet1") = Format(Request("HomeStreet1"))
	rsMember("HomeStreet2") = Format(Request("HomeStreet2"))
	rsMember("HomeCity") = Format(Request("HomeCity"))
	rsMember("HomeState") = Format(Request("HomeState"))
	rsMember("HomeZip") = Format(Request("HomeZip"))
	rsMember("HomeCountry") = Format(Request("HomeCountry"))
	rsMember("HomePhone") = Format(Request("HomePhone"))
	rsMember("SecondaryDescription") = Format( Request("SecondaryDescription") )
	rsMember("SecondaryStreet1") = Format(Request("SecondaryStreet1"))
	rsMember("SecondaryStreet2") = Format(Request("SecondaryStreet2"))
	rsMember("SecondaryCity") = Format(Request("SecondaryCity"))
	rsMember("SecondaryState") = Format(Request("SecondaryState"))
	rsMember("SecondaryZip") = Format(Request("SecondaryZip"))
	rsMember("SecondaryCountry") = Format(Request("SecondaryCountry"))
	rsMember("SecondaryPhone") = Format(Request("SecondaryPhone"))
	rsMember("SecondaryPExt") = Format(Request("SecondaryPExt"))

	rsMember.Update
End Sub


%>




<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
