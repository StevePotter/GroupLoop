<%
'-----------------------Begin Code----------------------------
'Log in a new home page hit
if not Request("Action") = "Old" then IncrementStat "HomePageHits"

Query = "SELECT Date, Body FROM News WHERE CustomerID = " & CustomerID & " ORDER BY Date DESC"
Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

if CalendarShowBirthdays = 1 and NewsShowEvents = 1 then
	Query = "Select NickName FROM Members WHERE ( CustomerID = " & CustomerID & " AND " & _
			"( Day(Birthdate) ='" & Day(Date) & "' AND Month(Birthdate) = '" & Month(Date) & "') )"
else
	Query = "Select NickName FROM Members WHERE ID = 0"
end if

Set rsBDays = Server.CreateObject("ADODB.Recordset")
rsBDays.Open Query, Connect, adOpenStatic, adLockReadOnly

if NewsShowEvents = 1 then
	Query = "SELECT Subject FROM Calendar WHERE (CustomerID = " & CustomerID & " AND StartDate <='" & Date & "' AND EndDate >='" & Date & "')"
else
	Query = "Select NickName FROM Members WHERE ID = 0"
end if

Set rsCalendar = Server.CreateObject("ADODB.Recordset")
rsCalendar.Open Query, Connect, adOpenStatic, adLockReadOnly

if not (rsPage.EOF AND rsBDays.EOF AND rsCalendar.EOF) then
	if not rsPage.EOF then
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="index.asp">
		<input type="hidden" name="Action" value="Old">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
	end if
'-----------------------End Code----------------------------
%>
	<p class="Heading" align="<%=HeadingAlignment%>"><%=NewsTitle%></p>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">News</td>
	</tr>
<%
	do until rsBDays.EOF
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMain %>" align="center"><font size=+1><b>Today</b></font></td>
				<td class="<% PrintTDMainSwitch %>"><font size=+1><b>Happy birthday, <%=rsBDays("Nickname")%>!</b></font></td>
			</tr>
<%
'-----------------------Begin Code----------------------------
		rsBDays.MoveNext
	loop
	do until rsCalendar.EOF
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMain %>" align="center"><b>Today</b></td>
				<td class="<% PrintTDMainSwitch %>"><b><a href="calendar_read.asp?ID=<%=FormatDateTime(Date, 2)%>"><%=rsCalendar("Subject")%></a></b></td>
			</tr>
<%
'-----------------------Begin Code----------------------------
		rsCalendar.MoveNext
	loop

	for i = 1 to rsPage.PageSize
		if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMain %>" align="center"><%=FormatDateTime(rsPage("Date"), 2)%></td>
					<td class="<% PrintTDMainSwitch %>"><%=rsPage("Body")%></td>
				</tr>
<%
'-----------------------Begin Code----------------------------
			rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
end if

set rsPage = Nothing
set rsBDays = Nothing
set rsCalendar = Nothing

'------------------------End Code-----------------------------
%>
<br>

<% PrintSource "HomeSource" %>

<br>

<!-- #include file="additions.asp" -->
