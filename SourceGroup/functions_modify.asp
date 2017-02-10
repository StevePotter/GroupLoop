<%
Function AddSubmitRequirement( strFieldName, strFieldDisplay )
	strReq = "if (form." & strFieldName & ".value == " & Chr(34) & Chr(34) & ")" & vbCrLf & _
			"strError += " & Chr(34) & "          You forgot the " & strFieldDisplay & ". \n" & Chr(34) & ";" & vbCrLf
	AddSubmitRequirement = strReq
End Function

Function IncludePrivacy( strSectionName )
	Query = "SELECT IncludePrivacy" & strSectionName & " FROM Look WHERE CustomerID = " & CustomerID
	Set rsPrivacy = Server.CreateObject("ADODB.Recordset")
	rsPrivacy.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
		DisplayPrivacy = CBool(rsPrivacy("IncludePrivacy" & strSectionName)) and not cBool(SiteMembersOnly)
	rsPrivacy.Close
	Set rsPrivacy = Nothing

	IncludePrivacy = DisplayPrivacy
End Function


'-------------------------------------------------------------
'This sub prints a field for editing an object
' RequiredField - is this necesary? if so, an * is shown
' strFieldLabel - the text displayed next to the actual field value
' strFieldName - the name of the field, both for the form and the recordset
' strFieldType - type of field - text, date, textarea
' intWidth - width of field
' intHeight - height of field
' rsObject - the recordset object containing the item
' strDisplaySecurity - who can see this?  admins or members?
' blConditional - if this is false, don't print the sub.  this is usually true, but can be false (in the case of displaypricay)
' this displays nothing if there is just a single item
' strRequirements - the string representing all the javascript requirements for submittal
'-------------------------------------------------------------
Sub PrintItemField( RequiredField, strFieldLabel, strFieldName, strFieldType, intWidth, intHeight, rsObject, strDisplaySecurity, blConditional, strRequirements )

		if not blConditional then Exit Sub

		if intWidth = "" or intWidth = 0 then intWidth = 50	'go with the default width
		if intHeight = "" or intHeight = 0 then intHeight = 20	'go with the default Height
		if strDisplaySecurity = "" then strDisplaySecurity = "Members"

		'This field is for admins only
		if (strDisplaySecurity = "Administrators" or strDisplaySecurity = "Admin") and not LoggedAdmin() then Exit Sub

		strFieldType = LCase(strFieldType)

		strReq = ""
		if RequiredField then
			strReq = "*&nbsp;"
			strRequirements = strRequirements & Add3Requirement( strFieldName, strFieldLabel )
		end if
%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right"><%=strReq & strFieldLabel%></td>
      	<td class="<% PrintTDMain %>">
<%
			if IsNull(rsObject) then
				strFieldData = ""
			else
				strFieldData = rsObject(strFieldName)
			end if

			if strFieldType = "text" then
%>
       		<input type="text" name="<%=strFieldName%>" size="<%=intWidth%>" value="<%=FormatEdit( strFieldData )%>">
<%
			elseif strFieldType = "checkbox" then
				PrintCheckBox strFieldData, strFieldName
			elseif strFieldType = "textarea" then
				TextArea strFieldName, intWidth, intHeight, True, ExtractInserts( strFieldData )
			elseif strFieldType = "datetime" or strFieldType = "date" then
				intShowTime = 0
				if strFieldType = "DateTime" then intShowTime = 1

				DatePulldown strFieldName, strFieldData, intShowTime
			end if
%>

     	</td>
    </tr>
<%
End Sub


Sub SubmitPage( strRequirements, strAction )
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
<%
			Response.Write strRequirements
%>
			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can <%=strAction%> this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</script>
<%
End Sub


Sub PrintItemForm( blOpenItem, strTable, strModSource, strRequestID, strKeyField, strCustomQuery, strSubmit, blShowFormatting )

	if blOpenItem then
		if Request(strRequestID) = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the " & strRequestID & "."))
		intID = CInt(Request(strRequestID))

		strMatch = "MemberID = " & Session("MemberID")
		if blLoggedAdmin then strMatch = "CustomerID = " & CustomerID

		if strCustomQuery = "" then
			Query = "SELECT * FROM " & strTable & " WHERE " & strKeyField & " = " & intID & " AND " & strMatch
		else
			Query = strCustomQuery
		end if
		rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		if rsEdit.EOF then
			Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
		end if

	end if


	strRequirements = ""
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back</a></p>
<%
	if blShowFormatting then
		if not UseWYSIWYGEdit() then Response.Write "<a href='formatting_view.asp' target='_blank'>Click here</a> for formatting tips.<br>"
		Response.Write "<a href='inserts_view.asp?Table=InfoPages' target='_blank'>Click here</a> for page inserts.<br>"
	end if
%>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%><%=strModSource%>" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
<%
	if blOpenItem then
		Response.Write "<input type=hidden name='SourceID' value='" & rsEdit(strKeyField) & "'>"
		Response.Write "<input type=hidden name='" & rsEdit("CommonID") & "' value='1'>"
		if IsObject(rsEdit("CommonID")) then Response.Write "<input type=hidden name='CommonID' value='" & rsEdit("CommonID") & "'>"
	end if

	PrintTableHeader 0
	PrintFields strRequirements

	SubmitPage strRequirements, LCase(strSubmit)
%>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="<%=strSubmit%>">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------

End Sub

'-------------------------------------------------------------
'This sub lists the other linked sites that an item belongs to
' strNoun - announcement, calendar event, etc.  display mask
' rsItems - recordset of items, linked by commonId
' this displays nothing if there is just a single item
'-------------------------------------------------------------
Function ListDupes( strNoun, rsItems )
		if rsItems.RecordCount > 1 then
	%>

			<tr> 
				<td class="<% PrintTDMain %>" colspan="2">
				<p>This <%=LCase(strNoun)%> also resides on <%=rsItems.RecordCount - 1%> other site(s).  Simply check the boxes next to the sites that you want to update with these changes.</p>
				<input type="hidden" name="Edit" value="Multiple">
	<%
			do until rsItems.EOF
	%>
				&nbsp;&nbsp;<%	PrintCheckBox 0, rsItems("ID") %> <%=GetSiteTitle(rsItems("CustomerID"))%><br>

	<%
				rsItems.MoveNext
			loop
	%>
     			</td>
   			</tr>
	<%
		end if
	ListDupes = True
End Function

Sub DeleteItem( strTable, strNoun, strThisSource, strModSource )

		if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
		intID = CInt(Request("ID"))

		strMatch = "MemberID = " & Session("MemberID")
		if blLoggedAdmin then strMatch = "CustomerID = " & CustomerID

		Query = "SELECT ID, CustomerID FROM " & strTable & " WHERE CommonID = " & intID & " AND " & strMatch

		Set rsUpdate = Server.CreateObject("ADODB.Recordset")
		rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
		if rsUpdate.EOF then
			set rsUpdate = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
		end if


		if rsUpdate.RecordCount > 1 and Request("Delete") = "Multiple" then
			Query = "DELETE " & strTable & " WHERE ID = " & intID

			do until rsUpdate.EOF
				if Request(rsUpdate("ID")) = "1" then
					Query = Query & " OR ID = " & rsUpdate("ID")

					RevQuery = "DELETE Reviews WHERE TargetTable = '" & strTable & "' AND TargetID = " & rsUpdate("ID")
					Connect.Execute RevQuery, , adCmdText + adExecuteNoRecords
				end if

				rsUpdate.MoveNext
			loop

			Connect.Execute Query, , adCmdText + adExecuteNoRecords

		'This is a multi-site addition
		elseif rsUpdate.RecordCount > 1 then
	%>
			<p>This <%=strNoun%> also resides on <%=rsUpdate.RecordCount - 1%> other site(s).  Simply check the boxes next to the sites you want to delete this <%=strNoun%> from.</p>
			<form method="post" action="<%=strThisSource%>" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = true; return this.submitted">
			<input type="hidden" name="ID" value="<%=intID%>">
			<input type="hidden" name="Delete" value="Multiple">

	<%
			do until rsUpdate.EOF
	%>
				&nbsp;&nbsp;<%	PrintCheckBox 0, rsUpdate("ID") %> <%=GetSiteTitle(rsUpdate("CustomerID"))%><br>

	<%
				rsUpdate.MoveNext
			loop
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="Submit" value="Delete">
			</form>
	<%
		else

			rsUpdate.Delete
			rsUpdate.Update

			Query = "DELETE Reviews WHERE TargetTable = '" & strTable & "' AND TargetID = " & intID
			Connect.Execute Query, , adCmdText + adExecuteNoRecords
	%>
			<p>The <%=strNoun%> has been deleted. &nbsp;<a href="<%=strModSource%>">Modify another.</a></p>
	<%
		end if

		rsUpdate.Close
	Set rsUpdate = Nothing

End Sub


Sub UpdateItem()
	if Request("SourceID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intSourceID = CInt(Request("SourceID"))
	intCommonID = Request("CommonID")

	'if they are a regular member, double check the memberID on the item
	strMatch = "AND MemberID = " & Session("MemberID")
	if blLoggedAdmin then strMatch = "AND CustomerID = " & CustomerID

	if intCommonID = "" then
		Query = "SELECT * FROM " & strTable & " WHERE ID = " & intSourceID & strMatch 
	else
		Query = "SELECT * FROM " & strTable & " WHERE CommonID = " & intCommonID & strMatch 
	end if
	rsEdit.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if


	do until rsEdit.EOF
		UpdateItemFields
		rsEdit.MoveNext
	loop

'------------------------End Code-----------------------------
%>
	<p>The <%=LCase(strNoun)%> has been edited. <br>
	<a href="<%=strViewSource & intSourceID%>"><%=strViewAction%> the updated <%=LCase(strNoun)%>.</a><br>
	<a href="<%=strListSource%>">View all <%=LCase(strPluralNoun)%>.</a>
	</p>
<%
'-----------------------Begin Code----------------------------
End Sub


Sub GoModify()
%>
	<p align="<%=HeadingAlignment%>"><span class=Heading>Modify <%=strPluralNoun%></span><br>
	<span class=LinkText><a href="<%=strListSource%>">Back To <%=strListSourceName%></a></span></p>
<%
	strSubmit = Request("Submit")

	if strSubmit = "Update" then
		if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")

		UpdateItem

	elseif strSubmit = "Delete" then
		DeleteItem strTable, LCase(strNoun), strModSource, strListSource

	elseif strSubmit = "Edit" then
		PrintItemForm True, strTable, strModSource, "ID", "ID", "", "Update", True
	else
		Redirect(strListSource)
	end if
End Sub
%>