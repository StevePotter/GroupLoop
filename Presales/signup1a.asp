<!-- #include file="header.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="dsn.asp" -->
<% AddHit "signup1a.asp" %>
<!-- #include file="closedsn.asp" -->

<p class=Heading align=center>
Multi-Site Instructions
</p>
<form METHOD="post" ACTION="signup2.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
			if Request("ParentID") <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=Request("ParentID")%>">
<%
			end if
%>


The sign-up process for the multi-site version works the following way:<br>
1.  You will first create your home site, which all other sites will stem off of.<br>
2.  After you have created the home site, you will create each child site.<br>

<input type="submit" name="Submit" value="I Understand">

</form>
<p align=center>

</p>


<!-- #include file="footer.asp" -->