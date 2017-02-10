<!-- #include file="photos_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CatsExist then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but no photos can be added until the administrator creates a category."))
if not CBool( IncludePhotos ) or not CBool( PhotosMembers ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember then Redirect("members.asp?Source=members_photos_add.asp")
if not (LoggedAdmin or CBool( PhotosMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20

intID = Request("ID")
if intID <> "" then intID = CInt(intID)
'------------------------End Code-----------------------------
%>

<script language="JavaScript">
<!--
	function submit_page(form) {
		//Error message variable
		var strError = "";
		//They don't want a thumbnail but didn't put in a name
		if (form.Thumbnail[1].checked && form.Name.value == "")
			strError += "          If you choose not to have a thumbnail, you must enter a description. \n";
		if (form.PhotoFile.value == "")
			strError += "          You forgot to select a file to upload. \n";

		if(strError == "") {
			alert('Uploading files may take some time, so please be patient.  After pressing OK, your photo will upload.  Please do not constantly press the Add button.');
			return true;
		}
        else{
			strError = "Sorry, but you must go back and fix the following errors before you can add your photos: \n" + strError;
			alert (strError);
			return false;
		}   
	}

//-->
</SCRIPT>


<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Photo</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<p>Remember that only gif, bitmap (bmp), and jpeg files may be uploaded (tif files are not accepted, sorry).  Anything else will be rejected.</p>

<%
if SiteMembersOnly = 0 then
%>
<p><b>Please note:</b> if you are uploading this to a private section, we recommend <b>not</b> creating a 
thumbnail.  Although the photo can't be viewed elsewhere, the thumbnail can.</p>
<%
end if
%>


<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_photos_add_process.asp" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
<input type="hidden" name="Password" value="<%=Session("Password")%>">
<% PrintTableHeader 0 %>
<tr>
	<td class="<% PrintTDMain %>"  align="right">
		Category
	</td>
	<td class="<% PrintTDMain %>">
<%
'-----------------------Begin Code----------------------------
		PrintCategoryPullDown intID, 1, 0, 0, 1, "PhotoCategories", "ID", ""
'------------------------End Code-----------------------------
%>
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		* Description of Photo
	</td>
	<td class="<% PrintTDMain %>">
		<input type="text" size="50" name="Name">
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		Photo File
	</td>
	<td class="<% PrintTDMain %>">
		<input type="file" name="PhotoFile">
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		Should a thumbnail (mini-version of the picture) be created to give people a quick 
		preview of the photo before viewing it?
	</td>
	<td class="<% PrintTDMain %>">
		<% PrintRadio 1, "Thumbnail" %>
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		Should the photo be optimized for the best screen fit and download time? (recommended)
	</td>
	<td class="<% PrintTDMain %>">
		<% PrintRadio 1, "Optimize" %>
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" align="center" colspan="2">

		<input type="submit" name="Submit" value="Add"></td>
	</td>
</tr>
</table>
</form>
