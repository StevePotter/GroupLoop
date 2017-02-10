<!-- #include file="admin_functions.asp" -->


<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin then	Redirect("members.asp?Source=admin_layout_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change The Layout Of Your Site</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Make Changes" then
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

	if Request("TotalPageWidth") <> "" then rsLook("TotalPageWidth") = Request("TotalPageWidth") & Request("TotalPageWidth" & "Percent")
	if Request("TotalPageAlignment") <> "" then rsLook("TotalPageAlignment") = Request("TotalPageAlignment")
	if Request("TitleAlignment") <> "" then rsLook("TitleAlignment") = Request("TitleAlignment")
	if Request("TitleVAlignment") <> "" then rsLook("TitleVAlignment") = Request("TitleVAlignment")
	if Request("TitleSpace") <> "" then rsLook("TitleSpace") = Request("TitleSpace")

	if Request("TopMenuShare") <> "" then rsLook("TopMenuShare") = Request("TopMenuShare")
	if Request("TopMenuAboveTitle") <> "" then rsLook("TopMenuAboveTitle") = Request("TopMenuAboveTitle")

	if Request("TopMenuAlignment") <> "" then rsLook("TopMenuAlignment") = Request("TopMenuAlignment")
	if Request("LeftMenuAlignment") <> "" then rsLook("LeftMenuAlignment") = Request("LeftMenuAlignment")
	if Request("RightMenuAlignment") <> "" then rsLook("RightMenuAlignment") = Request("RightMenuAlignment")
	if Request("TopMenuVAlignment") <> "" then rsLook("TopMenuVAlignment") = Request("TopMenuVAlignment")
	if Request("LeftMenuVAlignment") <> "" then rsLook("LeftMenuVAlignment") = Request("LeftMenuVAlignment")
	if Request("RightMenuVAlignment") <> "" then rsLook("RightMenuVAlignment") = Request("RightMenuVAlignment")



	if Request("LeftMenuWidth") <> "" then rsLook("LeftMenuWidth") = Request("LeftMenuWidth") & Request("LeftMenuWidth" & "Percent")
	if Request("RightMenuWidth") <> "" then rsLook("RightMenuWidth") = Request("RightMenuWidth") & Request("RightMenuWidth" & "Percent")

	if Request("BodyAlignment") <> "" then rsLook("BodyAlignment") = Request("BodyAlignment")
	if Request("BodyVAlignment") <> "" then rsLook("BodyVAlignment") = Request("BodyVAlignment")

	if Request("ShowFooter") <> "" then rsLook("ShowFooter") = Request("ShowFooter")
	if Request("FooterAlignment") <> "" then rsLook("FooterAlignment") = Request("FooterAlignment")

	rsLook.Update
	set rsLook = Nothing


	Redirect("write_header_footer.asp?Source=admin_layout_edit.asp?Submit=Changed")

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The layout changes have been made.  You can &nbsp;<a href="admin_layout_edit.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>
<%
'-----------------------Begin Code---------------------b-------
else
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
%>

	<form METHOD="post" ACTION="admin_layout_edit.asp" name="myForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">

		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Total Page Width</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="left" colspan="2">
				Some people like to change the maximum width of the web page to a certain percentage or fixed-size of each 
				visitor's screen.  However, choose your sizes carefully because a page that is too small won't look good.  Most
				people leave this at 100% with and aligned to the center. 
			</td>
		</tr>	
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				The entire page takes up 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintWidth "TotalPageWidth" %> of each visitor's screen.
			</td>
		</tr>			
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align entire page at 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "TotalPageAlignment", "" %> of each visitor's screen.
			</td>
		</tr>
		</table>
		<br>

		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Page Title</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="left" colspan="2">
				The page title is at the top of each page. 
			</td>
		</tr>		
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align title at 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "TitleAlignment", "TitleVAlignment" %> of it's space on the top of each page..
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Should a blank line be placed below the title to separate it from the rest of the page?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio rsLook("TitleSpace"), "TitleSpace" %>
			</td>
		</tr>

		</table>
		<br>
<%

	blLeftMenu = SectorHasButtons("Left")
	blRightMenu = SectorHasButtons("Right")
	blTopMenu = SectorHasButtons("Top")


	blMenus = blLeftMenu or blRightMenu or blTopMenu
	blMultiMenus = ( (CInt(blLeftMenu) + CInt(blRightMenu) + CInt(blTopMenu)) > 1 )


	if blMenus then
%>
		<% PrintTableHeader 100 %>
		<% if blMultiMenus then %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Button Menus
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="left" colspan="2">
				For each button menu, you can change the alignment of its buttons and the width of the menu.
			</td>
		</tr>
		<%
		end if

		strDisp = ""
		if blTopMenu then strDisp = " and Buttons"

		%>
		<tr><td class="TDHeader" colspan="2" align="center">Title <%=strDisp%> Across the TOP of the Page</td></tr>
<%
			strChecked1 = ""
			strChecked2 = ""
			if rsLook("TopMenuShare") = 0 then
				strChecked1 = "checked"
			else
				strChecked2 = "checked"
			end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				You can choose whether or not other menus <br>
				will be placed below the top menu.  
				<br><a href="http://www.GroupLoop.com/homegroup/images/menu_format_options.gif"><%=PrintTDLink("Click here for a visual diagram.")%></a>
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<input type="radio" name="TopMenuShare" value="0" <%=strChecked1%>> Option 1. Everything, including other menus, is placed below the top menu.  <a href="http://www.GroupLoop.com/homegroup/images/menu_format_option1.gif">Visual diagram.</a><br> 
			<input type="radio" name="TopMenuShare" value="1" <%=strChecked2%>> Option 2. The other menus start at the same vertical position as the top menu.  <a href="http://www.GroupLoop.com/homegroup/images/menu_format_option2.gif">Visual diagram.</a>
			</td>
		</tr>
<%
		if blTopMenu then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Place the top menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
<%
				PrintRadioOption "TopMenuAboveTitle", 1, "Above the Title<br>", rsLook("TopMenuAboveTitle")
				PrintRadioOption "TopMenuAboveTitle", 0, "Below the Title", rsLook("TopMenuAboveTitle")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align each button 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "TopMenuAlignment", "TopMenuVAlignment" %> on the page.
			</td>
		</tr>

<%
		end if
%>
		</table>
		<br>

<%
		if blLeftMenu then
		%>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Menu Across the LEFT of the Page</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				This menu takes up 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintWidth "LeftMenuWidth" %> of the page.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align each button 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "LeftMenuAlignment", "LeftMenuVAlignment" %> on the page.
			</td>
		</tr>
		</table>
		<br>
<%
		end if
		if blRightMenu then
		%>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Menu Across the RIGHT of the Page</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 This menu takes up 
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				<% PrintWidth "RightMenuWidth" %> of the page.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align each button 
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				<% AlignmentPulldown "RightMenuAlignment", "RightMenuVAlignment" %> on the page.
			</td>
		</tr>
		</table>
		<br>
<%
		end if
%>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Page Body</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" colspan="2"  align="left">
				The body is the main informational part of the page, which changes all the time.  You can change the alignment of the 
				body, but we highly recommend keeping it to the "Left" horizontally and "Top" vertically.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align the body 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "BodyAlignment", "BodyVAlignment" %> on the page.
			</td>
		</tr>
		</table>
		<br>

<%
	end if
%>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader" colspan="2" align="center">Footer</td></tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" colspan="2"  align="left">
				The footer is the bunch of text links at the bottom of the page.  You can change the alignment of the footer, and even turn it completely off.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				 Align the footer 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% AlignmentPulldown "FooterAlignment", "" %> on the page.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Should a footer be displayed?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio rsLook("ShowFooter"), "ShowFooter" %>
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
	set rsLook = Nothing
end if





'The pulldown with the different alignments
Sub AlignmentPulldown( strSelect, strVSelect )
%>
	<select name="<%=strSelect%>">
<%
		WriteOption "left", "Left", rsLook(strSelect)
		WriteOption "center", "Center", rsLook(strSelect)
		WriteOption "right", "Right", rsLook(strSelect)
%>
	</select> 
<%
	if strVSelect <> "" then
%>	
		horizontally and 
		<select name="<%=strVSelect%>">
<%
			WriteOption "top", "Top", rsLook(strVSelect)
			WriteOption "middle", "Middle", rsLook(strVSelect)
			WriteOption "bottom", "Bottom", rsLook(strVSelect)
%>
		</select> vertically  
<%
	end if
End Sub

Sub PrintWidth( strSelect )

	strWidth = rsLook( strSelect )

	if InStr( strWidth, "%" ) then
		strPercent = "%"
		intWidth = cInt( Left( strWidth, Len( strWidth ) - 1 ) )
	else
		strPercent = ""
		intWidth = cInt( strWidth )
	end if
%>
	<input type="text" size="4" name="<%=strSelect%>" value="<%=intWidth%>">&nbsp;

	<select name="<%=strSelect%>Percent">
<%
		WriteOption "", "Pixels (Screen Dots)", strPercent
		WriteOption "%", "Percent", strPercent
%>
	</select>
<%
End Sub
%>


