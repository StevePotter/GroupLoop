<%
if not LoggedAdmin then Redirect("members.asp?Source=admin_pagehits.asp")


Set cmd = Server.CreateObject("ADODB.Command")
With cmd
	.ActiveConnection = Connect
	.CommandText = "GetHomePageStats"
	.CommandType = adCmdStoredProc

	.Parameters.Refresh

	.Parameters("@CustomerID") = CustomerID

	.Execute , , adExecuteNoRecords
	intTotalPageHits = .Parameters("@TotalPageHits")
	intDifferentPages = .Parameters("@DifferentPages")
	intDifferentVisitors = .Parameters("@DifferentVisitors")
	intDifferentSessions = .Parameters("@DifferentSessions")
	intHomePageHits = .Parameters("@HomePageHits")
	 
End With
Set cmd = Nothing

startDate = DateAdd( "m", -1, Date )
endDate = Date
if Request("startMonth") <> "" then startDate = AssembleDate("Start")
if Request("endMonth") <> "" then endDate = AssembleDate("End")

if Request("startDate") <> "" then startDate = Request("startDate")
if Request("endDate") <> "" then endDate = Request("endDate")


strDateSQL = " Date >= '" & startDate & " 12:00:01 am ' AND Date <= '" & endDate & " 11:59:59 pm' " 

%>
<form method="post" action="admin_pagehits.asp" name="MyForm">

<%
if Request("Page") = "" then
%>
	<p align="<%=HeadingAlignment%>"><span class=Heading>Site Stats</span><br>
	<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

	<p><span class="SubHeading"><u>Summary:</u></span><br>
	Total Page Hits: <%=intTotalPageHits%><br>
	Different Pages: <%=intDifferentPages%><br>
	</p>


	<span class="SubHeading">Most Popular Pages:</span><br>
	Hits from: <%	DatePulldown "Start", startDate, 0 %> to 
	<%	DatePulldown "End", endDate, 0 %> <input type="submit" name="submit" value="Go"><br>
	
	
	
	Order By: 
	<select name="OrderPopPages" onChange="this.form.submit();">
	<%
		strOrderPopPages = Request("OrderPopPages")
		if strOrderPopPages = "" then strOrderPopPages = "PageHits DESC"

		WriteOption "PageHits DESC", "Total Page Hits", strOrderPopPages
		WriteOption "DistinctIPHits DESC", "Hits From Different IPs", strOrderPopPages
		WriteOption "Page", "Page Name", strOrderPopPages
	%>
	</select><br>
	

	


	<p>Click on a page to get a detailed list of its hits.</p>

	<% PrintTableHeader 0 %>
	<tr>
		<td class="TDHeader">Page</td>
		<td class="TDHeader">Total Hits</td>
		<td class="TDHeader">Hits From Different IPs</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>

	<%
	Query = "SELECT Page, COUNT(*) As PageHits, COUNT( DISTINCT IP ) As DistinctIPHits  FROM Hits WHERE ( CustomerID = " & CustomerID & _
		" AND " & strDateSQL & ") GROUP BY Page ORDER BY " & strOrderPopPages

		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = 100
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set Page = rsPage("Page")
		Set PageHits = rsPage("PageHits")
		Set DistinctIPHits = rsPage("DistinctIPHits")

		do until rsPage.EOF
	%>
		<tr>
			<td class="<% PrintTDMain %>"><a href="admin_pagehits.asp?Page=<%=Server.URLEncode(Page) & "&StartDate=" & startDate & "&EndDate=" & endDate%>"><%=PrintTDLink(PrintStart(Page))%></a></td>
			<td class="<% PrintTDMain %>"><%=PageHits%></td>
			<td class="<% PrintTDMain %>"><%=DistinctIPHits%></td>
			<td class="<% PrintTDMain %>" align="center"><a href="<%=Page%>"><%=PrintTDLink("View")%></a></td>
		</tr>
	<%
			rsPage.MoveNext
		loop
	%>
	</table>
	<%=Query%>
<%
else
	Query = "SELECT * FROM Hits WHERE ( CustomerID = " & CustomerID & " AND Page = '" & Request("Page") & "' " &_
		" AND " & strDateSQL & ") ORDER BY DATE Desc"

		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = 100
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set Page = rsPage("Page")
		Set MemberID = rsPage("MemberID")
		Set ItemDate = rsPage("Date")

%>
	<p class="Heading">Details for <a href="<%=Request("Page")%>"><%=Request("Page")%></a></p>
	
	<a href="admin_pagehits.asp">Back</a>
	<% PrintTableHeader 0 %>

	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Member</td>
	</tr>

	<%

		do until rsPage.EOF
	%>
		<tr>
			<td class="<% PrintTDMain %>"><%=ItemDate%></a></td>
			<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		</tr>
	<%
			rsPage.MoveNext
		loop
	%>
	</table>

<%

end if
	rsPage.Close
%>
</form>

