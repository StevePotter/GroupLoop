<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_schemes.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------

%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Scheme Manager</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%

strSubmit = Request("Submit")

'-------------------------------------------------------------
'Main menu
'-------------------------------------------------------------
if strSubmit = "" then
%>
	<p>The Scheme Manager allows you to easily change the entire look of your site.  You can 
	save the look and/or graphics of your site for later use.  So go ahead, save your current scheme and mess around.  You 
	can always get it back later!
	</p> 

	<strong>What would you like to do?</strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_schemes.asp?Submit=Save">Save your current scheme</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_schemes.asp?Submit=LoadCustom">Load one your saved schemes</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_schemes.asp?Submit=Modify">Modify one your saved schemes</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_schemes.asp?Submit=LoadOurPage">Load a GroupLoop.com scheme</a><br>
<%

'-------------------------------------------------------------
'Saving the current scheme
'-------------------------------------------------------------
elseif strSubmit = "Save" then
	'Make sure they entered a name
	if Request("Type") <> "" AND Request("Name") = "" then Redirect("incomplete.asp")

	if Request("Type") = "" then
		PrintTableHeader 0
%>
		<script language="JavaScript">
		<!--
			function submit_page(form) {
				//Error message variable
				var strError = "";
				if (form.Name.value == "")
					strError += "Sorry, but you forgot the name. \n";

				if(strError == "") {
					return true;
				}
				else{
					alert (strError);
					return false;
				}   
			}

		//-->
		</script>
		* indicates required information<br>
		<form method="post" action="admin_schemes.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
		<tr><td class="<% PrintTDMain %>">* What name would you like for this scheme?  <input type="text" name="Name" size="20" ></td></tr>
<%
		'Get the size of the current images folder
		strImageFolder = GetPath ("images")
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		Set TestFolder = FileSystem.GetFolder( strImageFolder )
		dblSize = TestFolder.Size
		Set TestFolder = Nothing
		Set FileSystem = Nothing
		'If there are no current images, don't give them the option to save the graphics
		if dblSize > 0 then
%>
			<tr><td class="<% PrintTDMain %>"><input type="radio" name="Type" value="Look" checked> Save the look of the site</td></tr>
			<tr><td class="<% PrintTDMain %>"><input type="radio" name="Type" value="Graphics"> Save the graphics of the site</td></tr>
			<tr><td class="<% PrintTDMain %>"><input type="radio" name="Type" value="All"> Save both the look and the graphics of the site</td></tr>
<%		else	%>
			<input type="hidden" name="Type" value="Look">
<%		end if	%>
		<tr><td class="<% PrintTDMain %>" valign="top">Description of Scheme <textarea name="Description" cols="55" rows="5" wrap="PHYSICAL"></textarea></td></tr>
		<tr><td class="<% PrintTDMain %>" align=center><input type="submit" name="Submit" value="Save"></td></tr>
		</form>
		</table>
<%
	'Just save the look
	elseif Request("Type") = "Look" then
		SetLook( 0 )	'The 0 means enter a new scheme
%>
		<p>The look scheme has been saved as <%=Request("Name")%>.  <a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>

<%
	elseif Request("Type") = "Graphics" then
		intID = SetGraphics	'SetGraphics is a function, so we gotta set it, but the intID isn't used (may work otherwise, but I took too much C++)
%>
		<p>The look scheme has been saved as <%=Request("Name")%>.  <a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>

<%
	else
		intID = SetGraphics	'Set the graphics and get the new schemeID
		SetLook( intID )	'Set the look to the new scheme
%>
		<p>The look and graphics schemes have been saved as <%=Request("Name")%>.  <a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>
<%
	end if

'-------------------------------------------------------------
'List the schemes, whether they are saved or our pre-made ones
'-------------------------------------------------------------
elseif strSubmit = "LoadCustom" or strSubmit = "LoadOurPage" then
	'The customerID on the samples schemes is -1
	if strSubmit = "LoadCustom" then
		Query = "SELECT ID, Type, Name, Description FROM Schemes WHERE CustomerID = " & CustomerID & " ORDER BY Date DESC"
	else
		Query = "SELECT ID, Type, Name, Description FROM Schemes WHERE CustomerID = -1 ORDER BY Date DESC"
	end if
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.CacheSize = PageSize
	rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsUpdate.EOF then
%>
		<p>Sorry, but you don't have any schemes saved right now.  Try loading an GroupLoop.com scheme.  <a href="javascript:history.back(1)">Click here</a> 
		to go back.</p>
<%
	else
		Set ID = rsUpdate("ID")
		Set SType = rsUpdate("Type")
		Set Name = rsUpdate("Name")
		Set Description = rsUpdate("Description")

		if strSubmit = "LoadCustom" then
%>
		<p>If you saved both the look and graphics in a scheme and wish to load it, you must load the Look and Graphics 
		first.  The order you load them doesn't matter, and they should have the same name.</p>
<%
		end if

		PrintTableHeader 0
%>
		<tr>
<%			if strSubmit = "LoadOurPage" then	'Include the screenshots for our schemes%>
				<td class="TDHeader">&nbsp;</td>
<%			end if	%>
			<td class="TDHeader">Scheme Type</td>
			<td class="TDHeader">Name</td>
			<td class="TDHeader">Description</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		do until rsUpdate.EOF
				if SType = "All" then
					strType = "Look and Graphics"
				else
					strType = SType
				end if
%>
				<form METHOD="POST" ACTION="admin_schemes.asp"	name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
				<input type="hidden" name="ID" value="<%=ID%>">
				<tr>
<%					if strSubmit = "LoadOurPage" then	%>
						<td class="<% PrintTDMain %>"><img src="http://www.grouploop.com/schemeshots/<%=ID%>.jpg"></td>
<%					end if	%>
					<td class="<% PrintTDMain %>"><%=strType%></td>
					<td class="<% PrintTDMain %>"><%=Name %></td>
					<td class="<% PrintTDMain %>">&nbsp;<%=Description%></td>
					<td class="<% PrintTDMain %>"><input type="submit" name="Submit" value="Load Scheme"></td>
				</tr>
				</form>
<%
			rsUpdate.MoveNext
		loop
		rsUpdate.Close
		Response.Write "</table>"
	end if

	set rsUpdate = Nothing

'-------------------------------------------------------------
'Load a scheme
'-------------------------------------------------------------
elseif strSubmit = "Load Scheme" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Type, CustomerID FROM Schemes WHERE ID = " & intID & " AND (CustomerID = -1 OR CustomerID = " & CustomerID & ")"
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp")
	end if

	blGroupLoop = False
	if rsUpdate("CustomerID") = -1 then blGroupLoop = True

	if rsUpdate("Type") = "Look" then
		GetLook rsUpdate("ID")
		'Our schemes, delete the current graphics
		if rsUpdate("CustomerID") = -1 then GetGraphics rsUpdate("ID"), blGroupLoop
	elseif rsUpdate("Type") = "Graphics" then
		GetGraphics rsUpdate("ID"), blGroupLoop
	elseif rsUpdate("Type") = "All" then
		GetLook rsUpdate("ID")
		GetGraphics rsUpdate("ID"), blGroupLoop
	end if
%>
	<!-- #include file="write_constants.asp" -->
<%

	Set rsUpdate = Nothing
	Redirect("write_header_footer.asp?Source=admin_schemes.asp?Submit=Changed")

elseif strSubmit = "Changed" then
	'This is here so changes can be seen right away

%>
	<p>The scheme has been loaded. <a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>
<%


'-------------------------------------------------------------
'List the schemes to edit or delete
'-------------------------------------------------------------
elseif strSubmit = "Modify" then
	Query = "SELECT ID, Type, Name, Description FROM Schemes WHERE CustomerID = " & CustomerID & " ORDER BY Date DESC"
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.CacheSize = PageSize
	rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsUpdate.EOF then
		Set rsUpdate = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you don't have any schemes saved right now.  Try loading a GroupLoop.com scheme."))
	end if

	Set ID = rsUpdate("ID")
	Set SType = rsUpdate("Type")
	Set Name = rsUpdate("Name")
	Set Description = rsUpdate("Description")

	PrintTableHeader 0
%>
	<tr>
		<td class="TDHeader">Scheme Type</td>
		<td class="TDHeader">Name</td>
		<td class="TDHeader">Description</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	do until rsUpdate.EOF
%>
		<form METHOD="POST" ACTION="admin_schemes.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
		<input type="hidden" name="ID" value="<%=ID%>">
		<tr>
			<td class="<% PrintTDMain %>"><%=SType%></a></td>
			<td class="<% PrintTDMain %>"><%=Name %></td>
			<td class="<% PrintTDMain %>">&nbsp;<%=Description%></td>
			<td class="<% PrintTDMain %>">
				<input type="Submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete Scheme" onClick="DeleteBox('If you delete this scheme, there is no way to get it back.  Are you sure?', 'admin_schemes.asp?Submit=DeleteScheme&ID=<%=ID%>')">			
			</td>
		</tr>
		</form>
<%
		rsUpdate.MoveNext
	loop
	Response.Write "</table>"

	rsUpdate.Close
	set rsUpdate = Nothing


'-------------------------------------------------------------
'Delete on of their schemes
'-------------------------------------------------------------
elseif strSubmit = "DeleteScheme" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Type FROM Schemes WHERE ID = " & intID & " AND CustomerID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	'Delete the folder with the graphics
	if rsUpdate("Type") = "Graphics" or rsUpdate("Type") = "All" then
		strDeleteFolder = GetPath("schemes") & intID
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		if FileSystem.FolderExists(strDeleteFolder) then FileSystem.DeleteFolder strDeleteFolder
		Set FileSystem = Nothing
	end if

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing


	Query = "DELETE Look WHERE SchemeID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
%>
	<p>The scheme has been deleted. &nbsp;<a href="admin_schemes.asp?Submit=Modify">Click here</a> to modify another.<br>
	<a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>
<%


'-------------------------------------------------------------
'Edit one of their schemse
'-------------------------------------------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Name, Description FROM Schemes WHERE ID = " & intID & " AND CustomerID = " & CustomerID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if


%>
		<script language="JavaScript">
		<!--
			function submit_page(form) {
				//Error message variable
				var strError = "";
				if (form.Name.value == "")
					strError += "Sorry, but you forgot the name. \n";

				if(strError == "") {
					return true;
				}
				else{
					alert (strError);
					return false;
				}   
			}

		//-->
		</script>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_schemes.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=intID%>">
	<%PrintTableHeader 0%>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Name</td>
     	<td class="<% PrintTDMain %>"> 
      		<input type="text" name="Name" size="50" value="<%=FormatEdit( rsEdit("Name") )%>">
    	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Description of Scheme</td>
     	<td class="<% PrintTDMain %>"> 
      		<textarea name="Description" cols="55" rows="5" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Description") )%></textarea>
    	</td>
	</tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
	</tr>

  	</table>
	</form>
<%

	rsEdit.Close
	set rsEdit = Nothing

'-------------------------------------------------------------
'Update the scheme they just edited
'-------------------------------------------------------------
elseif strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Name") = "" then Redirect("incomplete.asp")

	Query = "SELECT Name, Description FROM Schemes WHERE ID = " & intID & " AND CustomerID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("Description") = Format( Request("Description") )

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

%>
	<p>The scheme has been edited. &nbsp;<a href="admin_schemes.asp?Submit=Modify">Click here</a> to modify another.<br>
	<a href="admin_schemes.asp">Click here</a> to return to the Scheme Manager.</p>
<%

end if



Sub GetGraphics( intID, blGroupLoop )
		'Get the folder paths
		if blGroupLoop then
			if ParentSiteExists() then
				strSchemeFolder = GetPath ("..\..\homegroup\schemes") & intID & "\"
			else
				strSchemeFolder = GetPath ("..\homegroup\schemes") & intID & "\"
			end if
		else
			strSchemeFolder = GetPath ("schemes") & intID & "\"
		end if
		strImageFolder = GetPath ("images")
		strRootFolder = GetPath ("")
		'Not make sure the scheme folder exists.  If not, the graphics scheme has been lost, and we have an error
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		if not FileSystem.FolderExists( strSchemeFolder ) then
			if not blGroupLoop then
				Set FileSystem = Nothing
				Redirect("error.asp?Message=" & Server.URLEncode("The graphics for this scheme no longer exist.  If you didn't do anything wrong, please e-mail us.  Please include the Scheme ID, which is " & intID &"."))
			end if
		else
			FileSystem.DeleteFile strImageFolder&"*.*"
			FileSystem.CopyFile strSchemeFolder&"*.*", strImageFolder
		end if
		Set FileSystem = Nothing

End Sub

Sub GetLook( intID )
		'delete their current non-custom buttons
		Query = "DELETE MenuButtons WHERE Custom = 0 AND CustomerID = " & CustomerID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		'Open a new scheme
		Set rsScheme = Server.CreateObject("ADODB.Recordset")

		Query = "SELECT * FROM MenuButtons WHERE CustomerID = " & CustomerID
		rsScheme.Open Query, Connect, adOpenStatic, adLockOptimistic

		Query = "SELECT * FROM MenuButtons WHERE Custom = 0 AND SchemeID = " & intID
		Set rsLook = Server.CreateObject("ADODB.Recordset")
		rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic
		do until rsLook.EOF
			rsScheme.AddNew
			rsScheme("CustomerID") = CustomerID
			rsScheme("Position") = rsLook("Position")
			rsScheme("Name") = rsLook("Name")
			rsScheme("Show") = rsLook("Show")
			rsScheme("Align") = rsLook("Align")
			rsScheme.Update
			rsLook.MoveNext
		loop
		rsScheme.Close
		rsLook.Close


		Query = "SELECT * FROM Look WHERE SchemeID = " & intID
		rsScheme.Open Query, Connect, adOpenStatic, adLockReadOnly



		'Open up the look recordset
		Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
		rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

		for i = 0 to rsLook.Fields.Count - 1
			strField = rsLook(i).Name

			'Don't include these here
			blExclude = cBool( Instr(strField, "ID") or Left(strField, 8) = "InfoText" or Left(strField, 7) = "Include" or Left(strField, 8) = "ListType" or Left(strField, 7) = "Display" )

			if not blExclude then rsLook(strField) = rsScheme(strField)
		next

		rsLook.Update

		set rsScheme = Nothing
		set rsLook = Nothing
End Sub






'-------------------------------------------------------------
'This function adds the graphics into a scheme
'-------------------------------------------------------------
Function SetGraphics()
		strImageFolder = GetPath ("images")
		strRootFolder = GetPath("")
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		Set TestFolder = FileSystem.GetFolder( strImageFolder )
		if TestFolder.Size = 0 then
			Response.Write "<p>Since you have no graphics in use, no graphics scheme will be saved.</p>"
			Set TestFolder = Nothing
			Set FileSystem = Nothing
			Exit Function
		end if

		'Open up the look recordset
		Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
		Set rsLook = Server.CreateObject("ADODB.Recordset")
		rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly

		Set cmdTemp = Server.CreateObject("ADODB.Command")
		With cmdTemp
			.ActiveConnection = Connect
			.CommandText = "AddScheme"
			.CommandType = adCmdStoredProc
			.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
			.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
			.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 400 )
			.Parameters("@CustomerID") = CustomerID
			.Parameters("@Name") = Format( Request("Name") )
			.Execute , , adExecuteNoRecords

			intID = .Parameters("@ItemID")

		End With
		Set cmdTemp = Nothing

		'Open a new scheme
		Query = "SELECT * FROM Schemes WHERE ID = " & intID
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
		'Update the fields
			rsNew("Type") = "Graphics"
			rsNew("Name") = Format( Request("Name") )
			rsNew("Description") = Format( Request("Description") )
			rsNew("CustomerID") = CustomerID

		rsNew.Update
		rsNew.Close
		set rsNew = Nothing
		set rsLook = Nothing


		strNewFolder = GetPath ("schemes") & intID & "\"
		FileSystem.CreateFolder strNewFolder
		FileSystem.CopyFile strImageFolder&"*.*", strNewFolder

		Set TestFolder = Nothing
		Set FileSystem = Nothing

		SetGraphics = intID
End Function



'-------------------------------------------------------------
'This function adds the look into a scheme
'if intID = 0, new scheme.  Otherwise, use a scheme already created
'-------------------------------------------------------------
Sub SetLook( intID )
		'Open up the look recordset
		Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
		Set rsLook = Server.CreateObject("ADODB.Recordset")
		rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly

		if intID = 0 then
			strType = "Look"
		else
			strType = "All"
		end if

		'Add the scheme if we must
		if intID = 0 then
			Set cmdTemp = Server.CreateObject("ADODB.Command")
			With cmdTemp
				.ActiveConnection = Connect
				.CommandText = "AddScheme"
				.CommandType = adCmdStoredProc
				.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
				.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
				.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 400 )
				.Parameters("@CustomerID") = CustomerID
				.Parameters("@Name") = Format( Request("Name") )
				.Execute , , adExecuteNoRecords

				intID = .Parameters("@ItemID")

			End With
			Set cmdTemp = Nothing
		end if


		'Open a new scheme
		Query = "SELECT * FROM MenuButtons WHERE Custom = 0 AND CustomerID = " & CustomerID
		Set rsButtons = Server.CreateObject("ADODB.Recordset")
		rsButtons.Open Query, Connect, adOpenStatic, adLockReadOnly

		Query = "SELECT * FROM MenuButtons WHERE Custom = 0 AND CustomerID = " & CustomerID
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
		do until rsButtons.EOF
			rsNew.AddNew
			rsNew("SchemeID") = intID
			rsNew("Position") = rsButtons("Position")
			rsNew("Name") = rsButtons("Name")
			rsNew("Show") = rsButtons("Show")
			rsNew("Align") = rsButtons("Align")
			rsNew.Update
			rsButtons.MoveNext
		loop
		rsButtons.Close
		rsNew.Close

		Set rsButtons = Nothing

		'Open a new scheme
		Query = "SELECT * FROM Schemes WHERE ID = " & intID
		rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
		'Update the fields
			'If it's new, then it's a look scheme. otherwise, it's an all scheme
			rsNew("Type") = strType
			rsNew("Name") = Format( Request("Name") )
			rsNew("Description") = Format( Request("Description") )
			rsNew("CustomerID") = CustomerID
		rsNew.Update
		rsNew.Close

		'Create the loop table for the scheme
		Query = "SELECT * FROM Look"
		rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
		rsNew.AddNew
		rsNew("SchemeID") = intID

		for i = 0 to rsLook.Fields.Count - 1
			strField = rsLook(i).Name

			'Don't include these here
			blExclude = cBool( Instr(strField, "ID") or Left(strField, 8) = "InfoText" or Left(strField, 7) = "Include" or Left(strField, 8) = "ListType" or Left(strField, 7) = "Display" )

			if not blExclude then rsNew(strField) = rsLook(strField)
		next
		rsNew.Update

		rsNew.Close
		set rsNew = Nothing

		rsLook.Close
		set rsLook = Nothing
End Sub

%>
