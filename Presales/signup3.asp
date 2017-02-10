<!-- #include file="header.asp" -->
<!-- #include file="dsn.asp" -->
<!-- #include file="functions.asp" -->
<% AddHit "signup3.asp" %>

<p class=Heading align=center>
Step 3. Choose Your Initial Look
</p>

Simply press the 'Use This Look' button next to the desired choice.  This isn't permanent, and you can always change it later!

<%
'We are creating a new child site.. secret!
if Request("ParentID") <> "" then intParentID = Request("ParentID")

if Request("Version") = "" or ( Request("Version") <> "Gold" and Request("Version") <> "Free" and Request("Version") <> "Parent"  ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

strType = Request("Version")


Query = "SELECT ID, Name, Description FROM Schemes WHERE CustomerID = -1 ORDER BY ID DESC"
Set rsSchemes = Server.CreateObject("ADODB.Recordset")
rsSchemes.CacheSize = 20
rsSchemes.Open Query, Connect, adOpenStatic, adLockReadOnly

Set ID = rsSchemes("ID")
Set Name = rsSchemes("Name")
Set Description = rsSchemes("Description")

intCounter = 1
do until rsSchemes.EOF
%>
	<table width="100%" border=0 cellspacing=3 cellpadding=3>
	<tr><td>
	<img src="../images/schemeshots/<%=ID%>.jpg">
	</td><td valign="middle" class="BodyText">
		<b><%=intCounter%>. &nbsp;Name:</b> <%=Name %><br>
		<b>Description:</b> <%=Description %><br>
		<div align="center">
		<form METHOD="POST" ACTION="signup4.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
			if intParentID <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
			end if
%>
			<input type="hidden" name="Version" value="<%=strType%>">
			<input type="hidden" name="ID" value="<%=ID%>">
			<input type="submit" name="Submit" value="Use This Look">
		</form>
		</div>
	</td></tr>
	</table>
<%
	intCounter = intCounter + 1
	rsSchemes.MoveNext
loop
rsSchemes.Close
Set rsSchemes = Nothing
%>

<!-- #include file="closedsn.asp" -->

<!-- #include file="footer.asp" -->