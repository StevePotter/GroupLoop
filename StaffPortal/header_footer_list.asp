<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'-----------------------Begin Code----------------------------
		Const PageSize = 40
		Query = "SELECT ID, SubDirectory, DomainName, FirstName, LastName FROM Customers WHERE Removed = 0 ORDER BY Date DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = 100
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="customer_delete.asp">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
%>
				<tr>
					<td class="TDHeader" align="center">&nbsp;</td>
					<td class="TDHeader">CustomerID</td>
					<td class="TDHeader">Subdirectory/Domain</td>
					<td class="TDHeader">Name</td>
				</tr>
<%
					for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
						if rsPage("SubDirectory") <> "" then
							strSub = rsPage("SubDirectory")
						else
							strSub = rsPage("DomainName")
						end if
'------------------------End Code-----------------------------
%>
						<form METHOD="POST" ACTION="customer_delete.asp">
						<input type="hidden" name="ID" value="<%=rsPage("ID")%>">
						<input type="hidden" name="Subdirectory" value="<%=rsPage("Subdirectory")%>">
						<tr>
							<td class="<% PrintTDMain %>" target="_blank" align="center"><a href="http://www.GroupLoop.com/<%=rsPage("Subdirectory")%>/write_header_footer.asp">Go</a></td>
							<td class="<% PrintTDMain %>"><%=rsPage("ID")%></td>
							<td class="<% PrintTDMain %>"><%=strSub%></td>
							<td class="<% PrintTDMain %>"><%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%></td>
							</td>
						</tr>
						</form>
<%
'-----------------------Begin Code----------------------------
						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->