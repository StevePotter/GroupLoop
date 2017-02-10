<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=maintenance.asp&ID=" & intID)


Set cmd = Server.CreateObject("ADODB.Command")
With cmd
	.ActiveConnection = Connect
	.CommandText = "GetHomePageStats"
	.CommandType = adCmdStoredProc

	.Parameters.Refresh

	.Execute , , adExecuteNoRecords
	intTotalPageHits = .Parameters("@TotalPageHits")
	intDifferentPages = .Parameters("@DifferentPages")
	intDifferentVisitors = .Parameters("@DifferentVisitors")
	intDifferentSessions = .Parameters("@DifferentSessions")
	intHomePageHits = .Parameters("@HomePageHits")
	 
End With
Set cmd = Nothing
%>
<form method="post" action="stats_homesite.asp" name="MyForm">

<p align="<%=HeadingAlignment%>"><span class=Heading>Home Site Stats</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<p><span class="SubHeading">Summary:</span><br>
Total Page Hits: <%=intTotalPageHits%><br>
Home Page Hits: <%=intHomePageHits%><br>
Different Pages Visited: <%=intDifferentPages%><br>
Different Sessions: <%=intDifferentSessions%><br>
Different IP Addresses: <%=intDifferentVisitors%>
</p>


<span class="SubHeading">Most Popular Pages:</span><br>
<select name="OrderPopPages" onChange="this.form.submit();">
<%
	strOrderPopPages = Request("OrderPopPages")
	if strOrderPopPages = "" then strOrderPopPages = "PageHits DESC"

	WriteOption "PageHits DESC", "Total Page Hits", strOrderPopPages
	WriteOption "DistinctIPHits DESC", "Hits From Different IPs", strOrderPopPages
	WriteOption "Page", "Page Name", strOrderPopPages
%>
</select>

<% PrintTableHeader 0 %>
<tr>
	<td class="TDHeader">Page</td>
	<td class="TDHeader">Total Hits</td>
	<td class="TDHeader">Hits From Different IPs</td>
	<td class="TDHeader">&nbsp;</td>
</tr>

<%
Query = "SELECT Page, COUNT(*) As PageHits, COUNT( DISTINCT IP ) As DistinctIPHits  FROM Hits " & _
	"GROUP BY Page ORDER BY " & strOrderPopPages

	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set Page = rsPage("Page")
	Set PageHits = rsPage("PageHits")
	Set DistinctIPHits = rsPage("DistinctIPHits")

	do until rsPage.EOF
%>
	<tr>
		<td class="<% PrintTDMain %>"><%=Page%></td>
		<td class="<% PrintTDMain %>"><%=PageHits%></td>
		<td class="<% PrintTDMain %>"><%=DistinctIPHits%></td>
		<%
		if Page <> "index.asp" then
			strDir = "/homegroup"
		else
			strDir = ""
		end if
		%>
		<td class="<% PrintTDMain %>" align="center"><a href="http://www.GroupLoop.com<%=strDir%>/<%=Page%>">View</a></td>
	</tr>
<%
		rsPage.MoveNext
	loop
%>
</table>



</form>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->