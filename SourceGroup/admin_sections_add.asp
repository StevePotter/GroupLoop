<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

	<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Section</span><br>
	<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span><br>
	</p>
<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_sections_add.asp")
Session.Timeout = 20

if Request("Submit") = "Next" then
	if Request("Title") = "" or Request("NounSingular") = "" or Request("NounPlural") = "" then Redirect("incomplete.asp")

	'Add the group to the database
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddSection"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh

		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Categorize") = Request("Categorize")
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@Title") = Format(Request("Title"))
		.Parameters("@NounSingular") = Format(Request("NounSingular"))
		.Parameters("@NounPlural") = Format(Request("NounPlural"))
		.Parameters("@SectionViewSecurity") = Request("SectionViewSecurity")
		.Parameters("@ModifySecurity") = Request("ModifySecurity")
		.Parameters("@IncludePrivacy") = Request("IncludePrivacy")
		.Parameters("@DisplaySearch") = Request("DisplaySearch")
		.Parameters("@DisplayDaysOld") = Request("DisplayDaysOld")
		.Parameters("@InfoText") = GetTextArea(Request("InfoText"))
		.Parameters("@ListType") = Request("ListType")
		.Parameters("@RateItems") = Request("RateItems")
		.Parameters("@ReviewItems") = Request("ReviewItems")
		.Parameters("@DisplayDateList") = Request("DisplayDateList")
		.Parameters("@DisplayAuthorList") = Request("DisplayAuthorList")
		.Parameters("@DisplayPrivacyList") = Request("DisplayPrivacyList")
		.Parameters("@DisplayDateItem") = Request("DisplayDateItem")
		.Parameters("@DisplayAuthorItem") = Request("DisplayAuthorItem")

		.Execute , , adExecuteNoRecords

		intSectionID = .Parameters("@SectionID")
	End With
	Set cmdTemp = Nothing


	Query = "SELECT * FROM Sections WHERE ID = " & intSectionID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsUpdate("ListType") = Request("ListType")
	rsUpdate("ItemsPerLine") = Request("ItemsPerLine")


	rsUpdate("DisplayTitleList") = Request("DisplayTitleList")
	rsUpdate("DisplayTitleItem") = Request("DisplayTitleItem")

	for i = 1 to 10
		rsUpdate("FieldName" & i) = Format(Request("FieldName" & i))
		rsUpdate("FieldType" & i) = Request("FieldType" & i)
	next

	rsUpdate.Update
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Title.value == "")
				strError += "          You forgot the title. \n";
			if (form.NounSingular.value == "")
				strError += "          You forgot the singular name. \n";
			if (form.NounPlural.value == "")
				strError += "          You forgot the plural name. \n";


			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can update this: \n" + strError;
				alert (strError);
				return false;
			}   
		}
	//-->
	</script>

	<p><b>Now we are going to finish up!</b><br>
	All you need to do now is give some more details about each field you entered.<br>

	<form method="post" action="admin_sections_add.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=intSectionID%>">
	<% PrintTableHeader 0 %>



	<%
	for i = 1 to 10
		PrintTableHeader 100
		strName = rsUpdate("FieldName" & i)
		strType = rsUpdate("FieldType" & i)
		if strName <> "" then
	%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<%=strName%>
				</td>
			</tr>
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">Is this a required field?  If not, people don't have to enter anything.</td>
      			<td class="<% PrintTDMain %>"> 
					<% PrintRadio 1, "RequireFieldInput"&i %>
	 			</td>
			</tr>

			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">Display this field
				</td>
      			<td class="<% PrintTDMain %>"> 
			<%	PrintCheckBoxNew 1, "DisplayListField"&i, "show('ListDisplay" & i & "')", "hide('ListDisplay" & i & "')"  %> When listing <%=Request("NounPlural")%><br>
			<%	PrintCheckBoxNew 1, "DisplayItemField"&i, "show('ItemDisplay" & i & "')", "hide('ItemDisplay" & i & "')" %> When viewing an individual <%=Request("NounSingular")%><br>
	 			</td>
			</tr>
<%
			if strType = "TextSingle" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you want something to automatically appear in the box when someone adds an <span id="Singular">item</span>, enter it here.  For example, if you are 
				entering Body Weight, enter "lbs."  This way the person adding the <span id="Singular">item</span> doesn't have to, which will make things quicker and neater.</td>
      			<td class="<% PrintTDMain %>"> 
    				<input type="text" name="FieldInputDefault<%=i%>" size="20">
	 			</td>
			</tr>
<%
			elseif strType = "TextBox" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you want something to automatically appear in the box when someone adds an <span id="Singular">item</span>, enter it here.  For example, if you are 
				entering instructions, write out Step 1, Step 2, and let the person fill in the details for each step.  This way the person adding the <span id="Singular">item</span> doesn't have to, which will make things quicker and neater.</td>
      			<td class="<% PrintTDMain %>"> 
					<% TextArea "FieldInputDefault"&i, 55, 10, True, "" %>
	 			</td>
			</tr>
<%
			elseif strType = "Date" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you want a date to already be selected when a person adds an <span id="Singular">item</span>, enter it here.  This can be helpful if most dates are 
				going to be around or on a certain one.  Regardless of what you enter here, the person can change the date.</td>
      			<td class="<% PrintTDMain %>"> 
					<% DatePulldown "FieldInputDefault"&i, "", 0 %>
	 			</td>
			</tr>
<%
			elseif strType = "Currency" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you want a certain currency to already be entered when someone adds an <span id="Singular">item</span>, such as "$10.00", enter it here.</td>
      			<td class="<% PrintTDMain %>"> 
    				<input type="text" name="FieldInputDefault<%=i%>" size="20" value="$0.00">
	 			</td>
			</tr>
<%
			elseif strType = "Link" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you want a link to automatically appear in the box when someone adds an <span id="Singular">item</span>, enter it here.  For example, if you are 
				entering an e-mail, write "@aol.com", or if you are linking to a web site, maybe enter 'www.GroupLoop.com'.</td>
      			<td class="<% PrintTDMain %>"> 
    				<input type="text" name="FieldInputDefault<%=i%>" size="20" value="">
	 			</td>
			</tr>
<%
			elseif strType = "Option" then
%>			
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">For the option list, you must enter the different options here.  You can add more later.  All you do is write each option out in the 
				box and press enter between options.  It is very important that you press enter after an option.  This will separate them into a list.<br>
				For example, if listing positions for lacrosse players, just enter<br>
				Goalie<br>
				Defense<br>
				Midfield<br>
				Attack<br>
				Simple!
				</td>
      			<td class="<% PrintTDMain %>"> 
					<textarea name="FieldInputDefault<%=i%>" cols="50" rows="4" wrap="PHYSICAL"></textarea>
	 			</td>
			</tr>
<%
			end if

			Response.Write "</table>"

			blAllowLink = CBool( strType = "TextSingle"  or strType = "TextBox"  or strType = "Date" or strType = "Currency"  or strType = "Photo"  or strType = "Option" )
			if blAllowLink then
%>
			<span id="ListDisplay<%=i%>" <%=GetDisplay(1)%>>
			<%	PrintTableHeader 100	%>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
					When Listing <%=Request("NounPlural")%>
				</td>
			</tr>
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">If you chose to display this when listing <%=Request("NounPlural")%>, should this field be a link to view it as an <span id="Singular"><%=lCase(Request("NounSingular"))%></span>?  This is 
				is the same thing as how the subject links to the individual story or announcement on your site.  If you chose to have it as a link, it will be underlined.  It is recommended 
				that only one field links to the <span id="Singular"><%=lCase(Request("NounSingular"))%></span>, not all of them.  That looks ugly and can confuse people. 
				</td>
      			<td class="<% PrintTDMain %>"> 
					<%	PrintCheckBox 0, "LinkToItemField"&i %> Link to the individual <%=Request("NounSingular")%><br>
	 			</td>
			</tr>
			</table>
			</span>

<%
			end if
%>


			<span id="ItemDisplay<%=i%>" <%=GetDisplay(1)%>>
			<%	PrintTableHeader 100	%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
					When Showing a <%=Request("NounSingular")%>
				</td>
			</tr>
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right" width="50%">When viewing an individual <%=Lcase(Request("NounSingular"))%>, you can change the way this field is displayed.  First decide 
				if the field's name (<%=strName%>) should be displayed next to its data.  Then choose how it is displayed and aligned.
				</td>
      			<td class="<% PrintTDMain %>"> 
					<%	PrintCheckBox 1, "ItemShowNameField"&i %> Display "<%=strName%>:" before the data.<br>
					Align the field to the 	<select name="ItemAlignmentField<%=i%>" size="1">
											<%
												WriteOption "left", "Left", ""
												WriteOption "center", "Center", ""
												WriteOption "right", "right", ""
											%>
												</select> of the page.<br>
					Display the field <select name="ItemFormatField<%=i%>" size="1">
											<%
												WriteOption "plain", "on just a plain line", "plain"
												WriteOption "bullet", "with a bullet", ""
												WriteOption "table", "in a table", ""
												WriteOption "paragraph", "on a line with spacing around it", ""
												WriteOption "numbered", "in a numbered list", ""
											%>
												</select>.<br>

	 			</td>
			</tr>
			</table>
			</span>

<%
		end if
	next
%>

		<%	PrintTableHeader 100	%>
		<tr>
  			<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Done">
   			</td>
		</tr>
	</table>
	</form>
<%
	rsUpdate.Close
	Set rsUpdate = Nothing

elseif Request("Submit") = "Done" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intSectionID = CInt(Request("ID"))

	Query = "SELECT * FROM Sections WHERE CustomerID = " & CustomerID & " AND ID = " & intSectionID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	for i = 1 to 10
		rsUpdate("FieldInputDefault" & i) = Format(Request("FieldInputDefault" & i))
		rsUpdate("DisplayListField" & i) = GetCheckedResult(Request("DisplayListField" & i))
		rsUpdate("DisplayItemField" & i) = GetCheckedResult(Request("DisplayItemField" & i))
		rsUpdate("LinkToItemField" & i) = GetCheckedResult(Request("LinkToItemField" & i))
		rsUpdate("RequireFieldInput" & i) = GetCheckedResult(Request("RequireFieldInput" & i))
		rsUpdate("ItemShowNameField" & i) = GetCheckedResult(Request("ItemShowNameField" & i))
		rsUpdate("ItemAlignmentField" & i) = Request("ItemAlignmentField" & i)
		rsUpdate("ItemFormatField" & i) = Request("ItemFormatField" & i)

	next
%>
<p>The section has been added. &nbsp;<a href="members_sectionitems_add.asp?ID=<%=intSectionID%>">Click here</a> to add <%=rsUpdate("NounPlural")%> to it.<br>
</p>
<%
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
else
	'------------------------End Code-----------------------------
	%>


	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Title.value == "")
				strError += "          You forgot the title. \n";
			if (form.NounSingular.value == "")
				strError += "          You forgot the singular name. \n";
			if (form.NounPlural.value == "")
				strError += "          You forgot the plural name. \n";


			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can update this: \n" + strError;
				alert (strError);
				return false;
			}   
		}
	//-->
	</SCRIPT>
	* indicates required information<br>
	<form method="post" action="admin_sections_add.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<span id="TitleFirstCap">Section Properties</span>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right" width="50%">
				* Section Title
			</td>
			<td class="<% PrintTDMain %>" align="left" width="40%">
				<input type="text" name="Title" value="" size="60"  onChange="changeWord('Title', this.value);">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Display the title
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox 1, "DisplayTitleList" %> When listing <span id="Plural">items</span><br>
			<%	PrintCheckBox 1, "DisplayTitleItem" %> When viewing an individual <span id="Singular">item</span><br>
			</td>
		</tr>


		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Please enter the name (singular and plural) of the <span id="Plural">items</span> in the list.  If you are making a recipe section, 
				you would make the singular name "recipe", plural "recipes".
			</td>
			<td class="<% PrintTDMain %>" align="left">
				* Singular name (ex - 'movie') <input type="text" name="NounSingular" value="" size="30"  onChange="changeWord('Singular', this.value);"><br>
				* Plural name (ex - 'movies') <input type="text" name="NounPlural" value="" size="30"  onChange="changeWord('Plural', this.value);">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				How should <span id="Plural">items</span> be listed?
			</td>
			<td class="<% PrintTDMain %>" align="left">
<%
				PrintRadioOptionNew "ListType", "Table", "In a Table<br>", "Table", "show('IncludePrivacy');"
				PrintRadioOptionNew "ListType", "Bulleted", "Bulleted List<br>", "", "hide('IncludePrivacy');"
				PrintRadioOptionNew "ListType", "Numbered", "Numbered List<br>", "", "hide('IncludePrivacy');"
				PrintRadioOptionNew "ListType", "Plain", "Plain, Unordered List<br>", "", "hide('IncludePrivacy');"
%>
				<span id="IncludePrivacy" <%=GetDisplay(1)%>>
					How many <span id="Plural">items</span> should be displayed per line?  <input type="text" name="ItemsPerLine" value="1" size="4" >
				</span>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Should <span id="Plural">items</span> be categorized? 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio 0, "Categorize" %>
			</td>
		</tr>
	<%
		if not cBool(SiteMembersOnly) then
	%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can view the <span id="Title"></span>&nbsp;section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
	<%
				PrintRadioOptionNew "SectionViewSecurity", "Anyone", "Anyone<br>", "Anyone", "show('IncludePrivacy');"
				PrintRadioOptionNew "SectionViewSecurity", "Members", "Site Members Only<br>", "Anyone", "hide('IncludePrivacy');"
	'			PrintRadioOptionNew "SectionViewSecurity", "Administrators", "Only Administrators<br>", "Anyone", "hide('IncludePrivacy');"
	%>
			</td>
		</tr>
		</table>

		<span id="IncludePrivacy" <%=GetDisplay(1)%>>
		<%	PrintTableHeader 100	%>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right" width="40%">
				Can there be <span id="Plural">items</span> for members only?
			</td>
			<td class="<% PrintTDMain %>" align="left" width="60%">
	<%
				PrintRadioOption "IncludePrivacy", 1, "Yes, allow private <span id='Plural'>items</span> (public <span id='Plural'>items</span> can still be added)<br>", 1
				PrintRadioOption "IncludePrivacy", 0, "No, all <span id='Plural'>items</span> can be always read by anyone, even non-members<br>", 1
	%>
			</td>
		</tr>
		</table>
		</span>

		<%	PrintTableHeader 100	%>
	<%
		end if
	%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right" width="40%">
				Who can add <span id="Plural">items</span>?
			</td>
			<td class="<% PrintTDMain %>" align="left" width="60%">
	<%
				PrintRadioOption "ModifySecurity", "Members", "All Site Members<br>", "Members"
				PrintRadioOption "ModifySecurity", "Administrators", "Only Administrators<br>", "Members"
	%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Let people
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox 1, "RateItems" %> Rate <span id="Plural">items</span><br>
			<%	PrintCheckBox 1, "ReviewItems" %> Review <span id="Plural">items</span><br>
			<%	PrintCheckBox 1, "DisplaySearch" %> Search <span id="Plural">items</span><br>
			<%	PrintCheckBox 1, "DisplayDaysOld" %> View <span id="Plural">items</span> added in the last x number of days<br>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When listing <span id="Plural">items</span>, which information about each <span id="Singular">item</span> should be shown?
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox 1, "DisplayDateList" %> Date Written<br>
			<%	PrintCheckBox 1, "DisplayAuthorList" %> Author<br>
			<%  if not cBool(SiteMembersOnly) then %>
				<%	PrintCheckBox 1, "DisplayPrivacyList" %> Privacy (Public or Private)
			<%  end if %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
			When viewing an <b>individual</b> <span id="Singular">item</span>, which information should be shown?
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox 1, "DisplayDateItem" %> Date Written<br>
			<%	PrintCheckBox 1, "DisplayAuthorItem" %> Author<br>
			</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the section, please enter it here.</td>
			<td class="<% PrintTDMain %>"> 
			<% TextArea "InfoText", 55, 10, True, "" %>
			</td>
		</tr>




		<tr><td colspan=2 class=TDHeader align=center> <span id="SingularFirstCap">Item</span> Fields </td></tr>
		<tr><td colspan=2 class=<% PrintTDMain %> align=left>The <span id="Plural">items</span> you add to this section have different details that you must define.  
			For example, if you are listing favorite movies, you may want a field for when the movie came out, how many stars it got, the type of movie, who starred in it, a few screenshots of it, 
			and a description of it.  Doing that is very simple.<br><br>
			You can enter up to 10 different fields, but you don't have to.  Just leave any ones you don't want to use blank.  First you must enter the description of the field.  This would be 
			"Main Characters", "Stars", "Movie Description", etc. for the movies example.  Then you must enter the type of property each one is.  This is simple and there are several types:<br>
			<b>Single-Line Text</b> - This is simple, and is used for short fields.  The fields are entered in a familiar type of box, shown here: <input type="text" name="none" size="5"><br>
			<b>Big Text Box</b> - This is used for longer fields, such as a movie description or recipe directions.  The fields are entered in the same type of box as the one right above (the section message box).<br>
			<b>Date</b> - Select this if you want the field to be a date.<br>
			<b>Link</b> - This can be a link to a web site or someone's e-mail.<br>
			<b>Currency</b> - Select this for $$$ fields.<br>
			<b>Photo</b> - Use this if the field is a digital photo that may be uploaded.<br>
			<b>File</b> - Use this if the field is a regular file (text, Word, acrobat, etc).<br>
			<b>Member Pulldown</b> - This helps when you want to use a member's name as a field.  This is useful when a post involves another member.<br>
			<b>Option List</b> - This is a bit more advanced, but is very useful.  If you want to pre-define a list of possible field values for each <span id="Singular">item</span>, this is what you should select.  
			For example, if you are listing lacrosse players and want to include their position ('Attack', 'Middie', 'Defense', 'Goalie'), this is what you should choose.  It will force each <span id="Singular">item</span> 
			to use one thing from the list, which will save time by skipping typing and will avoid typos.
			<br>	
		</td>
		
		</tr>
	<%
		for i = 1 to 10
	%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					Field <%=i%>
				</td>
			</tr>
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">Field Description</td>
      			<td class="<% PrintTDMain %>"> 
    				<input type="text" name="FieldName<%=i%>" size="50">
	 			</td>
			</tr>
			<tr> 
	   			<td class="<% PrintTDMain %>" align="right">Field Type</td>
      			<td class="<% PrintTDMain %>"> 
					<% FieldTypePullDown "FieldType"&i, "TextSingle" %>
	 			</td>
			</tr>

	<%	next %>
		<tr>
  			<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Next">
   			</td>
		</tr>
	</table>
	</form>
<%
end if
%>

<%
Sub FieldTypePullDown( strName, strSelected )
%>
	<select name="<%=strName%>" size="1">
<%
	WriteOption "TextSingle", "Single-Line Text", strSelected
	WriteOption "TextBox", "Big Text Box", strSelected
	WriteOption "Date", "Date", strSelected
	WriteOption "Link", "Link", strSelected
	WriteOption "Currency", "Currency", strSelected
	WriteOption "Photo", "Photo", strSelected
	WriteOption "File", "File", strSelected
	WriteOption "MemberPulldown", "Member Pulldown", strSelected
	WriteOption "Option", "Option List", strSelected
%>
	</select>
<%
End Sub


'-------------------------------------------------------------
'This function prints yes and no radio boxes, highlighting the right one depending on the bool passed
'-------------------------------------------------------------
Sub PrintRadioOptionNew( Name, Value, Display, Selected, onClick )
	Response.Write "<input type='radio' name=" & Chr(34) & Name & Chr(34) & " value=" & Chr(34) & Value & Chr(34)
	if Value = Selected then Response.Write " checked"
	Response.Write " onClick=" & Chr(34) & onClick & Chr(34) & " > " & Display
End Sub

'-------------------------------------------------------------
'This function prints a check box, checking it depending on the bool passed
'-------------------------------------------------------------
Sub PrintCheckBoxNew( intBool, strName, onCheck, onUnCheck )
	strChecked = ""
	if intBool = 1 then strChecked = "checked"
%>
		<input type="checkbox" name="<%=strName%>" value="1" <%=strChecked%> onClick="if (this.checked == true )<%=onCheck%>; else <%=onUnCheck%>;">
<%
End Sub

Function GetDisplay( blDisplay )
	blDisplay = CBool(blDisplay)
	if blDisplay then
		GetDisplay = ""
	else
		GetDisplay = " style=" & chr(34) & "display: none;" & chr(34)
	end if
End Function
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->