<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=members_images_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Edit A File</span><br>
<span class=LinkText><a href="<%=NonSecurePath%>members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
strFileName = Request("FileName")
strPath = Request("Path")
strSource= Request("Source")

if strFileName = "" or strPath = "" then Redirect("incomplete.asp")

if not ( strPath = "inserts" or strPath = "storeitems" or strPath = "posts" or strPath = "images" or strPath = "storegroups" or strPath = "media" or strPath = "photos" or strPath = "schemes" ) then Redirect("error.asp")

'Make sure the image exists
Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strFullFile = GetPath(strPath) & strFileName
	if not FileSystem.FileExists( strFullFile ) then
		Set FileSystem = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The file does not exist."))
	end if

Set FileSystem = Nothing

'------------------------End Code-----------------------------
%>
	<form enctype="multipart/form-data" method="post" action="members_files_modify_process.asp" onsubmit="if (this.submitted) return false; this.submitted = true; return true" name="MyForm">
	<input type="hidden" name="Path" value="<%=Request("Path")%>">
	<input type="hidden" name="Source" value="<%=strSource%>">
	<input type="hidden" name="FileName" value="<%=strFileName%>">

	<p>Here is the current file: <a href="<%=strPath%>/<%=strFileName%>"><%=strFileName%></a></p>

<%	PrintTableHeader 0	%>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Overwrite file.  If you with to overwrite this file with a new one, please select the file here.  If not, 
			just leave it blank.
		</td>
		<td class="<% PrintTDMain %>">
			<input type="file" name="File">
		</td>
	</tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
