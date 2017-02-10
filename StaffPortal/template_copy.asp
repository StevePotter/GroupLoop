<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p class=Heading align=center>Copy a Template File</p>
<%
'-----------------------Begin Code----------------------------
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=template_copy.asp&ID=" & intID)


Set FileSystem = CreateObject("Scripting.FileSystemObject")
strTemplateFolder = Server.MapPath("..\templategroup") & "\"
strTemplate2Folder = Server.MapPath("..\templategroup2") & "\"

if Request("Submit") = "Copy" then
	strFile = Request("File")

	if strFile = "" then Redirect("error.asp?Message=" & Server.URLEncode("No filename was passed."))

	Query = "SELECT ID, SubDirectory, Version FROM Customers WHERE Removed = 0 AND ( Version = 'Free' OR Version = 'Gold' OR Version = 'Parent' OR Version = 'Child') ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	strSourceFile = strTemplateFolder & strFile
	strSourceFile2 = strTemplate2Folder & strFile

	if not FileSystem.FileExists( strSourceFile ) then Redirect("error.asp?Message=" & Server.URLEncode("The file " & strSourceFile & " does not exist."))
	if not FileSystem.FileExists( strSourceFile2 ) then Redirect("error.asp?Message=" & Server.URLEncode("The file " & strSourceFile2 & " does not exist."))

%>
		<div ID="infoParent" NAME="infoParent" CLASS=parent>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('info'); return false" ID="infoIm">
		Results Summary</a>
		</td></tr></table>
		</div>
		<div ID="infoChild" NAME="infoChild" CLASS=child>
		<% PrintTableHeader 100 %>
		<tr><td class="TDMain1">
<%


	do until rsPage.EOF
		strSubDir = rsPage("SubDirectory")
		if strSubDir = "" then 
			Response.Write rsPage("ID") & " has no subdirectory on record<br>"
		else
			strFolder = Server.MapPath("..\" & strSubDir)

			if FileSystem.FolderExists(strFolder) then
				Response.Write "Copying file to " & strFolder & "<br>"
				if rsPage("Version") <> "Child" then
					FileSystem.CopyFile strSourceFile, strFolder & "\" & strFile
				else
					FileSystem.CopyFile strSourceFile2, strFolder & "\" & strFile
				end if
			else
				Response.Write rsPage("ID") & " has no subdirectory<br>"
			end if
		end if
		rsPage.MoveNext
	loop
	rsPage.Close
	Set rsPage = Nothing

'------------------------End Code-----------------------------
%>
		</td></tr></table>
	</div>
	<p>
	The file has been copied. &nbsp;<a href="template_copy.asp">Click here</a> to copy another.<br>
	</p>
<%
'-----------------------Begin Code----------------------------
else
	Set Template = FileSystem.GetFolder(strTemplateFolder)

	PrintTableHeader 0

	for each objFile in Template.Files
'------------------------End Code-----------------------------
%>
		<form METHOD="POST" ACTION="template_copy.asp">
		<input type="hidden" name="File" value="<%=objFile.Name%>">
		<tr>
			<td class="<% PrintTDMain %>"><%=objFile.Name%></td>
			<td class="<% PrintTDMainSwitch %>">
				<input type="Submit" name="Submit" value="Copy">
			</td>
		</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
	next

	Set Template = Nothing

	Response.Write "</table>"
end if

Set FileSystem = Nothing

'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->