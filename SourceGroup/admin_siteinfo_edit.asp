<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_siteinfo_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Site Info</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
blMultiSites = ParentSiteExists() or ChildSiteExists()

if Request("Submit") = "Update" then
	Query = "SELECT Description, Keywords, FooterSource FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

	Query = "SELECT Title, ShortTitle FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic


	rsLook("Description") = Format( Request("Description") )
	rsLook("Keywords") = Format( Request("Keywords") )
	rsLook("FooterSource") = GetTextArea( Request("FooterSource") )


	rsConfig.Update
	rsConfig.Close
	set rsConfig = Nothing

	rsLook.Update
	set rsLook = Nothing
%>
	<!-- #include file="write_constants.asp" -->
<%
	Redirect("write_header_footer.asp?Source=admin_siteinfo_edit.asp?Submit=Changed")
elseif Request("Submit") = "Changed" then
	'This is here so changes can be seen right away
'------------------------End Code-----------------------------
%>
		<p>The site info has been changed. &nbsp;<a href="admin_siteinfo_edit.asp">Click here</a> to change it again.</p>
<%
'-----------------------Begin Code----------------------------
else
	Query = "SELECT Description, Keywords, FooterSource FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly


	if blMultiSites then
		Query = "SELECT ShortTitle FROM Configuration WHERE CustomerID = " & CustomerID
		Set rsConfig = Server.CreateObject("ADODB.Recordset")
		rsConfig.Open Query, Connect, adOpenStatic, adLockReadOnly

	end if
'------------------------End Code-----------------------------
%>
	<form METHOD="post" ACTION="<%=SecurePath%>admin_siteinfo_edit.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Site Title
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="Title" value="<%=Title%>" size="50">
			</td>
		</tr>
<%
	if blMultiSites then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Shortened Title - Used for links on the menu.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="ShortTitle" value="<%=rsConfig("ShortTitle")%>" size="50">
			</td>
		</tr>
<%
	end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				A short description of your site.
			</td>
			<td class="<% PrintTDMain %>" align="left">
    			<textarea name="Description" cols="55" rows="4" wrap="PHYSICAL"><%=rsLook("Description") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				A list of keywords to describe your site.
			</td>
			<td class="<% PrintTDMain %>" align="left">
    			<textarea name="Keywords" cols="55" rows="4" wrap="PHYSICAL"><%=rsLook("Keywords") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What text should be shown in the footer?  This is could be contact information, or anything else 
				that deserves to be on the bottom of every page.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% TextArea "FooterSource", 55, 4, True, rsLook("FooterSource") %>
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
	set rsLook = Nothing

	if blMultiSites then
		Set rsConfig = Nothing
	end if
end if
'------------------------End Code-----------------------------
%>
