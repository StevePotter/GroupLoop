<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_petpeeves_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A Pet Peeve</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'Add the story
if Request("Submit") = "Add" then
	if Request("Subject") = "" then Redirect("incomplete.asp")
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM PetPeeves"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
	rsNew.AddNew
		if Request("Private") = "1" then 
			intPrivate = 1
		else
			intPrivate = 0
		end if
		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Private") = intPrivate
		rsNew("Subject") = Format( Request("Subject") )
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
	<p>Your pet peeve has been added. &nbsp;<a href="members_petpeeves_add.asp">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>
	<p>If you only want members to be able to read it, you should check the private box.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_petpeeves_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
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
      		<td class="<% PrintTDMain %>" align="right">* Pet Peeve</td>
      		<td class="<% PrintTDMain %>"> 
    			<textarea name="Subject" cols="55" rows="4" wrap="PHYSICAL"></textarea>
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Details (story behind it, whatever)</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Body" cols="55" rows="10" wrap="PHYSICAL"></textarea>
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