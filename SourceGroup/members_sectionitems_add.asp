<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember() and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember() then Redirect("members.asp?Source=members_sectionitems_add.asp?ID=" & Request("ID"))
Session.Timeout = 20


'List their sections
if Request("ID") = "" then



else

	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intSectionID = CInt(Request("ID"))


	Query = "SELECT * FROM Sections WHERE ID = " & intSectionID
	Set rsSection = Server.CreateObject("ADODB.Recordset")
	rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsSection.EOF then Redirect("error.asp")

	Noun = rsSection("NounSingular")
	PluralNoun = rsSection("NounPlural")

	if rsSection("ModifySecurity") = "Administrators" and not LoggedAdmin() then
		Set rsSection = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry" & GetNickNameSession() & ", but only administrators may add " & PluralNoun & "."))
	end if
%>
	<p align="<%=HeadingAlignment%>"><span class=Heading>Add <%=PrintAn(Noun)%>&nbsp;<%=PrintFirstCap(Noun)%></span><br>
	<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span><br>
	</p>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";


			if(strError == "") {
<%

	for i = 1 to 10
		if rsSection("FieldName"&i) <> "" and rsSection("RequireFieldInput"&i) = 1 then
%>
			if (form.Field<%=i%>.value == "")
				strError += "          You forgot to enter something for <%=rsSection("FieldName"&i)%>. \n";
<%
		end if
	next
%>

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

	* Indicates Required Info

	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_sectionitems_add_process.asp" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=intSectionID%>">

<%
	PrintTableHeader 100

	for FieldNums = 1 to 10
		PrintInput FieldNums

	next

%>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Add">
    	</td>
    </tr>
	</table>
	</form>
<%

end if









Sub PrintInput( intFieldNum )
	'Make sure this is a valid field
	if rsSection("FieldName"&intFieldNum) <> "" and rsSection("FieldType"&intFieldNum) <> "" then
%>		<tr><td class="<% PrintTDMain %>" align="right" valign="middle">
<%
		if rsSection("RequireFieldInput"&intFieldNum) = 1 then Response.Write "* "


		FieldName = rsSection("FieldName"&intFieldNum)
		FieldType = rsSection("FieldType"&intFieldNum)

		Response.Write FieldName

		if FieldType = "Link" then
			Response.Write ". Address of link.  If linking to a web site, keep the 'http://'. For example, to link to Yahoo, make sure you enter 'http://www.Yahoo.com'.  " & _
			" If you are linking to an e-mail address, just enter it."
		elseif FieldType = "Photo" then
			Response.Write ". If you have a digital photo to add here, click Browse and select it."
		elseif FieldType = "File" then
			Response.Write ". If you have a file to add, click Browse and select it."
		end if
%>
		</td>
    	<td class="<% PrintTDMainSwitch %>"> 
<%
		FieldName = "Field" & intFieldNum
		Default = rsSection("FieldInputDefault" & intFieldNum)

		if FieldType = "TextSingle" or FieldType = "Link" or FieldType = "Currency" then
			Response.Write "<input type='text' name='" & FieldName & "' size=50 value='" & Default & "'>"
		elseif FieldType = "TextBox" then
			TextArea FieldName, 55, 15, True, Default
		elseif FieldType = "Date" then
%>		<% DatePulldown FieldName, Date, 0 %>&nbsp;&nbsp;<% PrintHours FieldName & "Hour" %><font size="+1">:</font> <% PrintMinutes FieldName & "Min" %> <% PrintHalf FieldName & "Half" %>
<%
		elseif FieldType = "Photo" or FieldType = "File" then
			Response.Write "<input type='file' name='" & FieldName & "' value='" & Default & "'>"
		elseif FieldType = "Option" then
			PrintPropertyPulldown FieldName, Default
		end if
%>

		</td>
		</tr>
<%
	end if

End Sub








Sub PrintHours( strName )
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"


	for i = 1 to 12
		Response.Write "<option value='" & i & "'>" & i & "</option>"
	next
	Response.Write "</select>"
End Sub

Sub PrintHalf( strName )
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"

		Response.Write "<option value='AM'>AM</option>"
		Response.Write "<option value='PM'>PM</option>"
	Response.Write "</select>"
End Sub

Sub PrintMinutes( strName)
	Response.Write "<select name='" & strName & "' size=1>"
	Response.Write "<option value=''> </option>"

	for i = 0 to 59
		stri = ""
		if i < 10 then stri = "0"
		Response.Write "<option value='" & i & "'>" & stri & i & "</option>"
	next
	Response.Write "</select>"
End Sub











Function PrintAn(strFollowWord)
	if IsNull( strFollowWord) then
		PrintAn = "a"
	elseif strFollowWord = "" then
		PrintAn = "a"
	else
		testchar = Lcase( Left( strFollowWord,1 ) )

		if Instr( testchar, "aeiou" ) then
			PrintAn = "an"
		else
			PrintAn = "a"
		end if
	end if
End Function


'-------------------------------------------------------------
'This sub takes the data from an item's field and parses it for a pulldown menu
'-------------------------------------------------------------
Sub PrintPropertyPulldown( strName, strOptionData )
	Dim OptionArray
	OptionArray = Split( strOptionData, "<br>", -1, 1 )
%>
	<select name="<%=strName%>" size="1">
<%
	for i = 0 to UBound( OptionArray )
		if OptionArray(i) <> "" then
%>
			<option value="<%=OptionArray(i)%>"><%=OptionArray(i)%></option>
<%
		end if
	next
%>
	</select>
<%
End Sub


Function PrintFirstCap( strWord )
	PrintFirstCap = UCase( Left(strWord, 1) ) & Right(strWord, Len(strWord)-1)

End Function
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->