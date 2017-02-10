<!-- #include file="header.asp" -->
<!-- #include file="..\dsn.asp" -->
<!-- #include file="..\functions.asp" -->
<!-- #include file="..\sourcegroup\functions.asp" -->

<p align="center"><span class=Heading>Add Account Options</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing your Customer ID.  Please go back to the Modify Account menu and use the links there."))
intCustomerID = CInt(Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Make Changes" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("Media") = "" then Redirect("incomplete.asp")
	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")
	strMedia = Request("Media")
	if strMedia = "Add" then
		intAllowMedia = 1
	else
		intAllowMedia = 0
	end if

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check the scheme to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "ValidCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@EMail") = strEMail
		.Parameters("@CCName") = strCCName
		.Parameters("@CCNumber") = strCCNumber
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The information you entered did not exactly match that of the account.  Remember that the credit card being billed for the account is the only one that will work.  Please try again."))
		end if

		'Get the subdirectory
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		strSubDir = .Parameters("@SubDirectory")
		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strVersion = .Parameters("@Version")
		if strVersion = "Free" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You must upgrade to the Gold Version before you can do this."))
		end if
		if strSubDir = "" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory was not in our records.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
		end if

		if intAllowMedia = 0 then
			'Check the folder.  better be there
			strFolder = Server.MapPath("../" & strSubDir & "/media")
			if not FileSystem.FolderExists(strFolder) then
				Set Command = Nothing
				Redirect("error.asp?Message=" & Server.URLEncode("Your media folder could not be found on our server.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
			end if
			'Just kill the folder then create it again.  easier...
			FileSystem.DeleteFolder strFolder
			FileSystem.CreateFolder strFolder
			Response.Write "<b>Removing...</b><br>1. Files Removed From Server<br>"

			'Now kill all the media entries
			.CommandText = "DeleteMedia"
			.Parameters.Refresh
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			Response.Write "2. Media Entries Removed From Database<br>"

		end if

		'Update the media settings
		.CommandText = "UpdateCustomerMedia"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@AllowMedia") = intAllowMedia
		.Execute , , adExecuteNoRecords

		if intAllowMedia = 0 then
			Response.Write "3. Media Permissions Removed From Site<br>"
		end if

	End With
	Set Command = Nothing
	Set FileSystem = Nothing
	strPath = Server.MapPath("../" & strSubDir) & "/"
'------------------------End Code-----------------------------
%>
	<!-- #include file="../sourcegroup/write_constants.asp" -->

	<p>
	Your options have been changed.  To complete the change <a href="http://www.GroupLoop.com/<%=strSubDir%>/write_header_footer.asp">click here</a>.  Thanks!
	</p>
<%
'-----------------------Begin Code---------------------------
else
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'See if they need the media section
		.ActiveConnection = Connect
		.CommandText = "GetCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID

		.Execute , , adExecuteNoRecords

		strVersion = .Parameters("@Version")
		if strVersion = "Free" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You must upgrade to the Gold Version before you can do this."))
		end if
		blAllowMedia = CBool(.Parameters("@AllowMedia"))
	End With

	Set Command = Nothing

%>
	<script language="JavaScript">
		function submit_page(form) {
			//Error message variable
			var strError = "";

			if(!form.Media.checked)
				strError += "You forgot to check the box.  If you don't, you can't make any changes.\n";

			if(strError == "") {
<%
			'We have to warn them if they are deleting the media section
			if blAllowMedia then
%>
				var where_to = confirm('If you remove your media section, all the files that have been uploaded will be deleted.  Are you sure?');
				if (where_to == true){
					return true;
				}
				else{
					return false;
				}   
<%
			else
%>
				return true;
<%
			end if
%>
			}
			else{
				alert (strError);
				return false;
			}   
		}

	</SCRIPT>

	<p>For changes to be made, we must validate your account information.  Please enter 
	your information <b>exactly</b> like you did when you signed up.  Otherwise, it won't work.</p>
	<form METHOD="POST" ACTION="account_options_edit.asp" name="Signup" onsubmit="return submit_page(this)">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<% PrintTableHeader 0 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Verify Account Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Account E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name on Credit Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Credit Card Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCNumber" size="20">
     		</td>
		</tr>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Options
     		</td>
		</tr>
<%
		if blAllowMedia then
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">You currently are signed up for the Media section.  Should we remove it?</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="checkbox" name="Media" value="Remove">
     		</td>
		</tr>
<%
		else
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">You currently are not signed up for the Media section.    This section allows members to upload 
			and share their favorite movies, sounds, etc.  Files are easily categorized and include descriptions.  
			And just like other additions, they can be rated and reviewed.  For an additional $5/month, you can 
			get the media section with 40 free megs of space.  This is our latest, hottest section!  Would you like to add it?</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="checkbox" name="Media" value="Add">
     		</td>
		</tr>
<%
		end if
%>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Make Changes">
	   		</td>
		</tr>
  	</table>
	</form>
<%
end if
%>

<!-- #include file="..\closedsn.asp" -->

<!-- #include file="footer.asp" -->