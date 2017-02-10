<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then	Redirect("members.asp?Source=admin_buttons_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Menu Button</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add Button" then
	intNumMenus = 1

	Query = "SELECT * FROM MenuButtons"
	Set rsButtons = Server.CreateObject("ADODB.Recordset")
	rsButtons.Open Query, Connect, adOpenStatic, adLockOptimistic


	rsButtons.AddNew

	rsButtons("CustomerID") = CustomerID
	rsButtons("Position") = 2
	rsButtons("Custom") = 1

	rsButtons("Align") = Request("Align")
	rsButtons("Show") = Request("Show" )
	rsButtons("CustomLink") = Request("CustomLink")
	rsButtons("CustomLabel") = Request("Label")
	rsButtons("Name") = Request("Label")


	rsButtons.Update
	rsButtons.Close
	Set rsButtons = Nothing

	Redirect("write_header_footer.asp?Source=admin_buttons_add.asp?Submit=Changed")

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The button changes have been made.  You can <a href="admin_buttons_modify.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>
<%
'-----------------------Begin Code----------------------------
else

%>

	<SCRIPT LANGUAGE="JavaScript">
	<!--


	//-->
	</SCRIPT>


	<p>You can easily change the order and placement of your buttons below.</p>
	<form METHOD="post" ACTION="admin_buttons_add.asp" name="myForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">



		Page Position:
		<select name="Align">
		<option value="Left" selected>Left</option><option value="Right">Right</option><option value="Top">Top</option>
		</select><br>


		Show Button In:
		<select name="Show">
		<option value="Menu">Main Menu Only</option><option value="Footer">Footer Only</option><option value="MenuFooter" selected>Main Menu And Footer</option><option value="Nowhere">Do Not Show Button</option>
		</select><br>


		Label:
		<input type="text" size="15" name="Label" value=""><br>

		Link Button to:
		<input type="text" size="15" name="CustomLink" value=""><br>

		<input type="submit" name="Submit" value="Add Button"><br>

	</form>
<%
'-----------------------Begin Code----------------------------
end if
%>
