<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin() then Redirect("members.asp?Source=admin_account_remove.asp")

Set Command = Server.CreateObject("ADODB.Command")

With Command
	'Check to make sure the CC info is correct
	.ActiveConnection = Connect
	.CommandText = "GetOwnerMemberID"
	.CommandType = adCmdStoredProc
	.Parameters.Refresh
	.Parameters("@CustomerID") = CustomerID
	.Execute , , adExecuteNoRecords
	intOwnerID = .Parameters("@MemberID")
	strName = .Parameters("@FirstName") & "&nbsp;" & .Parameters("@LastName")
End With
Set Command = Nothing

if Session("MemberID") <> intOwnerID then Redirect("message.asp?Source=members.asp&Message=" & Server.URLEncode("Sorry, but only the site owner, " & strName & " can terminate the site."))

%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Terminate Your Site</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			var Confirmation = confirm('If you remove your account, you can never get it back.  Are you completely sure?');
			if (Confirmation == true){
				return true;
			}
			else{
				return false;
			}  
		}

	//-->
	</SCRIPT>

	<p>Please remember that once you terminate your account, it is <b>permanently</b> deleted.  Be 
	completely sure before doing this.</p>

	<form METHOD="POST" ACTION="http://www.GroupLoop.com/admin/account_remove.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">

	<p align="center">
	<input type="hidden" name="CustomerID" value="<%=CustomerID%>">
	<input type="hidden" name="MemberID" value="<%=intOwnerID%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="submit" name="Submit" value="Remove My Account">
	</p>

	</form>
