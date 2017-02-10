<!-- #include file="header.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="dsn.asp" -->
<% AddHit "signup2.asp" %>
<!-- #include file="closedsn.asp" -->

<p class=Heading align=center>
Step 2. Read and Agree to the TOS
</p>


<!-- #include file="tossource.asp" -->

<p align=center>

<form METHOD="post" ACTION="signup3.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
'We are creating a new child site.. secret!
if Request("ParentID") <> "" then intParentID = Request("ParentID")

if Request("Submit") = "" or ( Request("Submit") <> "Use Gold Version" and Request("Submit") <> "Use Free Version" and Request("Submit") <> "I Understand" and Request("Submit") <> "Use Multi-Site Version" ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

if Request("Submit") = "Use Gold Version" then
	strType = "Gold"
elseif Request("Submit") = "I Understand" then
	strType = "Parent"
else
	strType = "Free"
end if
			if intParentID <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
			end if
%>
	<input type="hidden" name="Version" value="<%=strType%>">
	<input type="submit" value="I Have Read And Agreed To The TOS">
</form>

</p>

<!-- #include file="footer.asp" -->