<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin then	Redirect("members.asp?Source=admin_advanced_visuals_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change The Advanced Visuals</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
	strImagePath = GetPath("images")

if Request("Submit") = "Make Changes" then
	Query = "SELECT CellSpacing, CellPadding, Border, Title FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic

	rsConfig("CellSpacing") = Request("CellSpacing")
	rsConfig("CellPadding") = Request("CellPadding")
	rsConfig("Border") = Request("Border")

	rsConfig.Update
	rsConfig.Close

	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic

	rsConfig("CustomHeader") = Request("CustomHeader")
	if SectorHasButtons("Top") then SetMenuStuff "Top"
	if SectorHasButtons("Left") then SetMenuStuff "Left"
	if SectorHasButtons("Right") then SetMenuStuff "Right"

	rsConfig.Update
	rsConfig.Close
	Set rsConfig = Nothing


	Sub SetMenuStuff( strMenu )
		if ImageExists(strMenu & "MenuSeparatorImage", strExt) then
			rsConfig(strMenu & "MenuSeparatorHeaderText") = Request(strMenu & "MenuSeparatorHeaderText")
			rsConfig(strMenu & "MenuSeparatorFooterText") = Request(strMenu & "MenuSeparatorFooterText")
		else
			response.write strMenu & "MenuSeparator"
			rsConfig(strMenu & "MenuSeparator") = Request(strMenu & "MenuSeparator")
		end if

		rsConfig(strMenu & "MenuButtonHeaderText") = Request(strMenu & "MenuButtonHeaderText")
		rsConfig(strMenu & "MenuButtonFooterText") = Request(strMenu & "MenuButtonFooterText")

		if ImageExists(strMenu & "MenuBackgroundImage", strExt) then
			rsConfig(strMenu & "MenuFullBackground") = Request(strMenu & "MenuFullBackground")
		end if


	End Sub

%>
<!-- #include file="write_constants.asp" -->

<%
	Redirect("write_header_footer.asp?Source=admin_advanced_visuals_edit.asp?Submit=Changed")

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The changes have been made.  You can &nbsp;<a href="admin_advanced_visuals_edit.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>
<%
'-----------------------Begin Code---------------------b-------
else
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

%>

	<form METHOD="post" ACTION="admin_advanced_visuals_edit.asp" name="myForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">

		<% PrintTableHeader 100 %>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Tables (item listings, etc.  These options are in a table)
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Spacing (if you don't know, leave it)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="2" name="CellSpacing" value="<%=CellSpacing%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Padding
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="2" name="CellPadding" value="<%=CellPadding%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Borders
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="2" name="Border" value="<%=Border%>">
			</td>
		</tr>

		</table>
		<br>

		<% PrintTableHeader 100 %>

<%
	if SectorHasButtons("Top") then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Top Menu Visuals
			</td>
		</tr>
<%
		if ImageExists("TopMenuSeparatorImage", strExt) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text before and after the image that separates the buttons on the top menu.  These are usually blank in the top menu, but sometimes can be a line break or space.  
					You may enter anything you like.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Before separator <input type="text" size="5" name="TopMenuSeparatorHeaderText" value="<%=rsLook("TopMenuSeparatorHeaderText")%>"><br>
					After separator <input type="text" size="5" name="TopMenuSeparatorFooterText" value="<%=rsLook("TopMenuSeparatorFooterText")%>">
				</td>
			</tr>

<%
		else
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text separator between buttons on the top menu.  You can put an image separator in the <a href="admin_images_edit.asp"><%=PrintTDLink("graphics")%></a> section.  
					  This is usually a line break, space, bar separator (' | ' ), or nothing.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" size="5" name="TopMenuSeparator" value="<%=rsLook("TopMenuSeparator")%>">
				</td>
			</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Text before and after each button on the top menu.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Before button <input type="text" size="5" name="TopMenuButtonHeaderText" value="<%=rsLook("TopMenuButtonHeaderText")%>"><br>
				After button <input type="text" size="5" name="TopMenuButtonFooterText" value="<%=rsLook("TopMenuButtonFooterText")%>">
			</td>
		</tr>
<%
		if ImageExists("TopMenuBackgroundImage", strExt) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				How far down does the menu background span?
			</td>
			<td class="<% PrintTDMain %>" align="Top">
<%
				PrintRadioOption "TopMenuFullBackground", 1, "The full width of the page<br>", rsLook("TopMenuFullBackground")
				PrintRadioOption "TopMenuFullBackground", 0, "To the end of the buttons", rsLook("TopMenuFullBackground")
%>			
		
			</td>
		</tr>

<%
		end if

	end if
	if SectorHasButtons("Left") then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Left Menu Visuals
			</td>
		</tr>
<%
		if ImageExists("LeftMenuSeparatorImage", strExt) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text before and after the image that separates the buttons on the left menu.  These are usually line breaks ('&lt;br&gt;'). You may enter anything you like.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Before separator <input type="text" size="5" name="LeftMenuSeparatorHeaderText" value="<%=rsLook("LeftMenuSeparatorHeaderText")%>"><br>
					After separator <input type="text" size="5" name="LeftMenuSeparatorFooterText" value="<%=rsLook("LeftMenuSeparatorFooterText")%>">
				</td>
			</tr>

<%
		else
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text separator between buttons on the left menu.  You can put an image separator in the <a href="admin_images_edit.asp"><%=PrintTDLink("graphics")%></a> section.  
					  This is usually a line break ('&lt;br&gt;'), space, bar separator (' | ' ), or nothing.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" size="5" name="LeftMenuSeparator" value="<%=rsLook("LeftMenuSeparator")%>">
				</td>
			</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Text before and after each button on the left menu.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Before button <input type="text" size="5" name="LeftMenuButtonHeaderText" value="<%=rsLook("LeftMenuButtonHeaderText")%>"><br>
				After button <input type="text" size="5" name="LeftMenuButtonFooterText" value="<%=rsLook("LeftMenuButtonFooterText")%>">
			</td>
		</tr>
<%
		if ImageExists("LeftMenuBackgroundImage", strExt) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				How far down does the menu background span?
			</td>
			<td class="<% PrintTDMain %>" align="Left">
<%
				PrintRadioOption "LeftMenuFullBackground", 1, "The full height of the page<br>", rsLook("LeftMenuFullBackground")
				PrintRadioOption "LeftMenuFullBackground", 0, "To the end of the buttons", rsLook("LeftMenuFullBackground")
%>			
		
			</td>
		</tr>

<%
		end if

	end if

	if SectorHasButtons("Right") then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Right Menu Visuals
			</td>
		</tr>
<%
		if ImageExists("RightMenuSeparatorImage", strExt) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text before and after the image that separates the buttons on the right menu.  These are usually line breaks ('&lt;br&gt;'). You may enter anything you like.
				</td>
				<td class="<% PrintTDMain %>" align="right">
					Before separator <input type="text" size="5" name="RightMenuSeparatorHeaderText" value="<%=rsLook("RightMenuSeparatorHeaderText")%>"><br>
					After separator <input type="text" size="5" name="RightMenuSeparatorFooterText" value="<%=rsLook("RightMenuSeparatorFooterText")%>">
				</td>
			</tr>

<%
		else
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Text separator between buttons on the right menu.  You can put an image separator in the <a href="admin_images_edit.asp"><%=PrintTDLink("graphics")%></a> section.  
					  This is usually a line break ('&lt;br&gt;'), space, bar separator (' | ' ), or nothing.
				</td>
				<td class="<% PrintTDMain %>" align="right">
					<input type="text" size="5" name="RightMenuSeparator" value="<%=rsLook("RightMenuSeparator")%>">
				</td>
			</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Text before and after each button on the right menu.
			</td>
			<td class="<% PrintTDMain %>" align="right">
				Before button <input type="text" size="5" name="RightMenuButtonHeaderText" value="<%=rsLook("RightMenuButtonHeaderText")%>"><br>
				After button <input type="text" size="5" name="RightMenuButtonFooterText" value="<%=rsLook("RightMenuButtonFooterText")%>">
			</td>
		</tr>
<%
		if ImageExists("RightMenuBackgroundImage", strExt) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				How far down does the menu background span?
			</td>
			<td class="<% PrintTDMain %>" align="Right">
<%
				PrintRadioOption "RightMenuFullBackground", 1, "The full height of the page<br>", rsLook("RightMenuFullBackground")
				PrintRadioOption "RightMenuFullBackground", 0, "To the end of the buttons", rsLook("RightMenuFullBackground")
%>			
		
			</td>
		</tr>

<%
		end if

	end if

%>
		</table>
		<br>



		<% PrintTableHeader 100 %>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Custom Header
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Are you using a custom header?  This should only Yes if GroupLoop has put a custom header in place for you.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio rsLook("CustomHeader"), "CustomHeader" %>
			</td>
		</tr>
		</table>
		<br>

		<% PrintTableHeader 100 %>
		<tr><td class="<% PrintTDMain %>" colspan="2" align="center"><input type="submit" name="Submit" value="Make Changes"></td></tr>
		</table>

	</form>
<%
'-----------------------Begin Code----------------------------
	rsLook.Close
	Set rsLook = Nothing

end if

%>


