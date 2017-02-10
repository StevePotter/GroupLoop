<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeStories ) then Redirect("error.asp")
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_dickmoves_add.asp")

Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A Dick Move</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'-------------------------------------------------------------
'This function writes a pulldown menu for members
'-------------------------------------------------------------
Sub PrintMemberPullDown( intMemberID )
	intMemberID = CInt(intMemberID)
	'Now we are going to get the group names to list in the pull-down menu
	Query = "SELECT ID, NickName FROM Members WHERE (CustomerID = " & CustomerID & ")"
	Set rsTempMembers = Server.CreateObject("ADODB.Recordset")
	rsTempMembers.Open Query, Connect, adOpenStatic, adLockReadOnly
	
	'Make the size 3 if there are many members
	if rsTempMembers.RecordCount <= 30 then
		%><select name="MemberID" size="1"><%
	else
		%><select name="MemberID" size="3"><%
	end if

	'We have passed a 0, which is non member
	if intMemberID = 0 then Response.Write "<option value = '0' SELECTED>Non-Member</option>" & vbCrlf

	do While not rsTempMembers.EOF
		'Highlight the current section
		if rsTempMembers("ID") = intMemberID then
			Response.Write "<option value = '" & rsTempMembers("ID") & "' SELECTED>" & rsTempMembers("NickName") & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & rsTempMembers("ID") & "'>" & rsTempMembers("NickName") & "</option>" & vbCrlf
		end if
		rsTempMembers.MoveNext
	loop
	rsTempMembers.Close
	set rsTempMembers = Nothing
	Response.Write("</select>")
End Sub


'Add the story
if Request("Submit") = "Add" then
	if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT ID, Private, MemberID, Subject, Body, CustomerID, IP, ModifiedID FROM Stories"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
	rsNew.AddNew
		if Request("Private") = "1" then 
			rsNew("Private") = 1
		else
			rsNew("Private") = 0
		end if
		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Subject") = Format( Request("Subject") )
		rsNew("TargetID") = Request("MemberID")
		rsNew("Points") = Request("Points")
		rsNew("Body") = Format( Request("Body") )
		rsNew("CustomerID") = CustomerID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
	rsNew.Update
	rsNew.MoveNext
	rsNew.MovePrevious
	intID = rsNew("ID")
	rsNew.Close
	Set rsNew = Nothing
'------------------------End Code-----------------------------
%>
	<p>That fucking dick is on record now, baby. &nbsp;<a href="members_dickmoves_add.asp">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------

else
	if not LoggedMember then Redirect("members.asp?Source=members_dickmoves_add.asp")

'------------------------End Code-----------------------------
%>
	<p>If you only want members to be able to read it, you should check the private box.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_dickmoves_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Private?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Private" value="1">
			</td>
   		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Dick Points</td>
      		<td class="<% PrintTDMain %>"> 
				<select name="Points">
<%
				for i = 1 to 10
					%><option value="<%=i%>"><%=i%></option><%
				next
%>
				</select>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Dick</td>
      		<td class="<% PrintTDMain %>"> 
       			<% PrintMemberPullDown -1 %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Story</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Body" cols="55" rows="20" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if
'------------------------End Code-----------------------------
%>