<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin() then Redirect("members.asp?Source=admin_space_add.asp")

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

if Version <> "Gold" then Redirect("error.asp")

strSubmit = Request("Submit")
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add Disk Space</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%


if strSubmit = "Update My Account" then
	intNewPhotosMegs = CInt(Request("PhotoMegs"))
	intNewMediaMegs = CInt(Request("MediaMegs"))


	Query = "SELECT PhotosMegs, MediaMegs FROM Configuration WHERE CustomerID = " & CustomerID
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
	<!-- #include file="write_constants.asp" -->
	<p>
	The space has been added.  Thank you!
	</p>
<%
'-----------------------Begin Code---------------------------
else
	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	Set CheckFolder = FileSystem.GetFolder( Server.MapPath("photos" ))
	dblPhotosSize = Round(CheckFolder.Size / 1000000, 1)
	Set CheckFolder = Nothing

	Set CheckFolder = FileSystem.GetFolder( Server.MapPath("media" ))
	dblMediaSize = Round(CheckFolder.Size / 1000000, 1)
	Set CheckFolder = Nothing

	Set FileSystem = Nothing

%>
	<script language="JavaScript">
	<!--
		//Throw out all the stuff we don't want ($)
		function ConvertInteger(currCheck) {
			if (!currCheck) return '';
			for (var i=0, currOutput='', valid="0123456789"; i<currCheck.length; i++)
				if (valid.indexOf(currCheck.charAt(i)) != -1)
					currOutput += currCheck.charAt(i);
			return currOutput;
		}

		function submit_page(form) {
			form.PhotoMegs.value = ConvertInteger(form.PhotoMegs.value)
			form.MediaMegs.value = ConvertInteger(form.MediaMegs.value)

			if (form.PhotoMegs.value == "")
				form.PhotoMegs.value == "0";
			if (form.MediaMegs.value == "")
				form.MediaMegs.value == "0";

			if (form.PhotoMegs.value == "0" && form.MediaMegs.value == "0"){
				strError = "Sorry, but you must enter an amount to add.";
				alert (strError);
				return false;
			}
			



			return true;

		}
	//-->
	</SCRIPT>
	<form METHOD="post" ACTION="admin_space_add.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
	<tr>
    	<td class="TDHeader" colspan=2 align="center"> 
    		Photos Section
    	</td>
	</tr>
	<tr>
    	<td class="<% PrintTDMain %>" colspan=2 align="left"> 
    		You currently have <%=PhotosMegs%> megs for for photos.  <%=dblPhotosSize%> megs are used 
			(<%=Round((PhotosMegs - dblPhotosSize), 1)%> megs available for new photos).  You may purchase additional 
			space for your photos at <b>$0.50 per meg, per month</b>.  To give you a better picture of how much space you may need, 
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
   			You currently have <%=MediaMegs%> megs for for files.  <%=dblMediaSize%> megs are used 
			(<%=Round((MediaMegs - dblMediaSize), 1)%> megs available for new files).  You may purchase additional 
			space for your files at <b>$0.50 per meg, per month</b>.  There is no way to know what people will upload, so we can't 
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
end if
%>
