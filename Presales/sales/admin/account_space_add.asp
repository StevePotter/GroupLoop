<!-- #include file="header.asp" -->
<!-- #include file="..\dsn.asp" -->
<!-- #include file="..\functions.asp" -->

<p align="center"><span class=Heading>Add Disk Space</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing your Customer ID.  Please go back to the Modify Account menu and use the links there."))
intCustomerID = CInt(Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Update My Account" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" then Redirect("incomplete.asp")

	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")

	intNewPhotosMegs = CInt(Request("PhotoMegs"))
	intNewMediaMegs = CInt(Request("MediaMegs"))


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

	End With
	Set Command = Nothing

	Query = "SELECT PhotosMegs, MediaMegs FROM Configuration WHERE CustomerID = " & intCustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		Set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("PhotosMegs") = rsUpdate("PhotosMegs") + intNewPhotosMegs
	rsUpdate("MediaMegs") = rsUpdate("MediaMegs") + intNewMediaMegs

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>
	Your account has been updated.  To return to your site, <a href="http://www.GroupLoop.com/<%=Request("Subdirectory")%>">click here</a>.  Thanks!
	</p>
<%
'-----------------------Begin Code---------------------------
elseif strSubmit = "Add Space" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" then Redirect("incomplete.asp")
	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check to make sure the CC info is correct
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

		'Get the customer's shit
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		strVersion = .Parameters("@Version")
		if strVersion = "Free" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You must upgrade to the Gold Version before you can do this."))
		end if
		strSubdirectory = .Parameters("@Subdirectory")

		.CommandText = "GetConfigDiskSpace"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		intPhotosMegs  = .Parameters("@PhotosMegs")
		intMediaMegs  = .Parameters("@MediaMegs")
	End With
	Set Command = Nothing

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	Set CheckFolder = FileSystem.GetFolder( Server.MapPath("..\" & strSubdirectory & "\photos" ))
	dblPhotosSize = Round(CheckFolder.Size / 1000000, 1)
	Set CheckFolder = Nothing

	Set CheckFolder = FileSystem.GetFolder( Server.MapPath("..\" & strSubdirectory & "\media" ))
	dblMediaSize = Round(CheckFolder.Size / 1000000, 1)
	Set CheckFolder = Nothing

	Set FileSystem = Nothing

%>
	<form METHOD="POST" ACTION="account_space_add.asp">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<input type="hidden" name="EMail" value="<%=strEMail%>">
	<input type="hidden" name="CCName" value="<%=strCCName%>">
	<input type="hidden" name="CCNumber" value="<%=strCCNumber%>">
	<input type="hidden" name="Subdirectory" value="<%=strSubdirectory%>">
	<% PrintTableHeader 0 %>
	<tr>
    	<td class="TDHeader" colspan=2 align="center"> 
    		Photos Section
    	</td>
	</tr>
	<tr>
    	<td class="<% PrintTDMain %>" colspan=2 align="left"> 
    		You currently have <%=intPhotosMegs%> megs for for photos.  <%=dblPhotosSize%> megs are used 
			(<%=Round((intPhotosMegs - dblPhotosSize), 1)%> megs available for new photos).  You may purchase additional 
			space for your photos at <b>$0.80 per meg, per month</b>.  To give you a better picture of how much space you may need, 
			1 meg usually holds about 15 photos.
    	</td>
	</tr>
	<tr> 
   		<td class="<% PrintTDMain %>" align="right">How many more megs would you like?</td>
   		<td class="<% PrintTDMain %>"> 
			<input type="text" name="PhotoMegs" size="4" value="0">
		</td>
	</tr>
	<tr>
    	<td class="TDHeader" colspan=2 align="center"> 
    		Media Section
    	</td>
	</tr>
	<tr>
   		<td class="<% PrintTDMain %>" colspan=2 align="left"> 
   			You currently have <%=intMediaMegs%> megs for for files.  <%=dblMediaSize%> megs are used 
			(<%=Round((intMediaMegs - dblMediaSize), 1)%> megs available for new photos).  You may purchase additional 
			space for your photos at <b>$0.60 per meg, per month</b>.  There is no way to know what people will upload, so we can't 
			tell you how many megs to add.  However, people usually add about 20 megs at a time.
   		</td>
	</tr>
	<tr> 
  		<td class="<% PrintTDMain %>" align="right">How many more megs would you like?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="MediaMegs" size="4" value="0">
		</td>
	</tr>
	<tr>
   		<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update My Account">
   		</td>
	</tr>
  	</table>
	</form>
<%
'Get their info
else
%>
	<p>Before we can edit your account, we must validate your account information.  Please enter 
	your information <b>exactly</b> like you did when you signed up.  Otherwise, it won't work.</p>
	<form METHOD="POST" ACTION="account_space_add.asp">
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
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add Space">
	   		</td>
		</tr>
  	</table>
	</form>
<%
end if

Function GetSelected( strComp1, strComp2 )
	if strComp1 = strComp2 then
		GetSelected = " selected"
	else
		GetSelected = ""
	end if
End Function
%>

<!-- #include file="..\closedsn.asp" -->

<!-- #include file="footer.asp" -->