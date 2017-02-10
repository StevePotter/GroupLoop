<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_account_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Your Account</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<p>
Only the owner of the account (not just an administrator) can make account changes, since they are  
responsible for paying for it.  Simply click on any of the links below to be taken to a secure page where 
the account changes will take place.
</p>

<a href="https://www.OurClubPage.com/<%=SubDirectory%>/admin_account_edit.asp?MemberID=<%=Session("MemberID")%>&Password=<%=Session("Password")%>">Change Account/Billing Information</a><br>
<a href="admin_space_add.asp">Add Space for Photos or Media (if you have it) Sections</a><br>
<a href="https://www.ourclubpage.com/admin/account_custom_add.asp?CustomerID=<%=CustomerID%>">Request Custom Graphics/Sections</a><br>
<a href="admin_account_remove.asp">Terminate Your Site</a>
<%
	if Version = "Parent" or Version = "Child" then
%>
		<br><a href="http://www.GroupLoop.com/homegroup/signup1.asp?MemberID=<%=Session("MemberID")%>&ParentID=<%=CustomerID%>">Add a Sub-Site</a><br>
<%
	end if
%>
