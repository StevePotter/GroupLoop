<!-- #include file="calendar_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeCalendar ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=CalendarTitle%></span><br>
<%
if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( CalendarMembers )) then
%>
<span class=LinkText><a href="members_calendar_add.asp">Add an Event</a></span>
<%
end if
%>
</p>

<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")

intRateCalendar = RateCalendar
intReviewCalendar = ReviewCalendar

Set rsList = Server.CreateObject("ADODB.Recordset")

'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Query = "SELECT ID, MemberID, Subject, Body FROM Calendar WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	rsList.CacheSize = 100
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		Set MemberID = rsList("MemberID")
		Set Body = rsList("Body")
		Set Subject = rsList("Subject")
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if


Public ListTypeCalendar, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT IncludePrivacyCalendar, InfoTextCalendar, DisplaySearchCalendar, DisplayDateListCalendar, DisplayAuthorListCalendar, DisplayPrivacyListCalendar  FROM Look WHERE CustomerID = " & CustomerID
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	InfoText = rsList("InfoTextCalendar")
	DisplaySearch = CBool(rsList("DisplaySearchCalendar"))
	DisplayDate = CBool(rsList("DisplayDateListCalendar"))
	DisplayAuthor = CBool(rsList("DisplayAuthorListCalendar"))
	DisplayPrivacy = (CBool(rsList("DisplayPrivacyListCalendar")) and CBool(rsList("IncludePrivacyCalendar"))) and not cBool(SiteMembersOnly)
rsList.Close

if DisplaySearch then
%>
<form METHOD="POST" ACTION="calendar.asp">
	Search For <input type="text" name="Keywords" size="25">
	<input type="submit" name="Submit" value="Go"><br>
</form>
<%
end if

if intSearchID <> "" then
	'Their search came up empty
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="calendar.asp">click here</a> to view all events.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="calendar.asp">click here</a> to view all events.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
		rsPage.CacheSize = PageSize
		Set TargetID = rsPage("TargetID")
%>
		<form METHOD="POST" ACTION="calendar.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Start Date</td>
			<td class="TDHeader">End Date</td>
			<% if DisplayAuthor then %>
			<td class="TDHeader">Author</td>
			<% end if %>	
			<td class="TDHeader">Subject</td>
			<% if intRateCalendar = 1  and intReviewCalendar = 0 then %>
				<td class="TDHeader" align=center>Rating</td>
			<% elseif intRateCalendar = 0  and intReviewCalendar = 1 then %>
				<td class="TDHeader" align=center>Review</td>
			<% elseif intRateCalendar = 1  and intReviewCalendar = 1 then %>
				<td class="TDHeader" align=center>Rating</td>
			<% end if %>	

			<% if DisplayPrivacy then %>
			<td class="TDHeader">Public?</td>
			<% end if %>	
		</tr>
<%
		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT ID, StartDate, EndDate, Date, MemberID, Subject, TotalRating, TimesRated, Private FROM Calendar WHERE CustomerID = " & CustomerID
		rsList.CacheSize = PageSize
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set StartDate = rsList("StartDate")
		Set EndDate = rsList("EndDate")
		Set MemberID = rsList("MemberID")
		Set TotalRating = rsList("TotalRating")
		Set TimesRated = rsList("TimesRated")
		Set Subject = rsList("Subject")
		Set IsPrivate = rsList("Private")

		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				rsList.Filter = "ID = " & TargetID
			%>
				<tr>
					<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(StartDate, 2)%></td>
					<td class="<% PrintTDMain %>" align="center"><%=FormatDateTime(EndDate, 2)%></td>
					<% if DisplayAuthor then %>
					<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
					<% end if %>	
					<td class="<% PrintTDMain %>"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink( Subject )%></a></td>
			<%		if intRateCalendar = 1 and intReviewCalendar = 0 then
			%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
						<font size="-2"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
			<%		elseif intRateCalendar = 0 and intReviewCalendar = 1 then
						if ReviewsExist( "Calendar", ID ) then
			%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Read/Add Review ")%></a></font></td>
			<%			else
			%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Add Review ")%></a></font></td>
			<%			end if
					elseif intRateCalendar = 1 and intReviewCalendar = 1 then
						if ReviewsExist( "Calendar", ID ) then
			%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
								<font size="-2"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate and Read/Add Review ")%></a></font></td>
			<%			else
			%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
							<font size="-2"><a href="calendar_event_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate/Add Review ")%></a></font></td>
			<%			end if
					end if%>
					<% if DisplayPrivacy then %>
					<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
					<% end if %>	
				</tr>
<%
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
		set rsPage = Nothing
		set rsList = Nothing
	end if
'They are just cycling through the events.  No searching.
else
	if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"

	' Constants for the days of the week
	Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7

	' Check for valid month input and set the current month
	if IsEmpty(Request("month")) OR NOT IsNumeric(Request("month")) then
	  datToday = Date()
	  intThisMonth = Month(datToday)
	elseif CInt(Request("month")) < 1 OR CInt(Request("month")) > 12 then
	  datToday = Date()
	  intThisMonth = Month(datToday)
	else
	  intThisMonth = CInt(Request("month"))
	end if

	' Check for valid year input and set the current year
	if IsEmpty(Request("year")) OR NOT IsNumeric(Request("year")) then
	  datToday = Date()
	  intThisYear = Year(datToday)
	else
	  intThisYear = CInt(Request("year"))
	end if

	'Set the month name and 
	strMonthName = MonthName(intThisMonth)
	'Sets a date of the first day of the month
	datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
	'Gets the day of the week of the first day in the month
	intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
	'Gets the last day of the month
	intLastDay = GetLastDay(intThisMonth, intThisYear)
		
	' Get the previous month and year
	intPrevMonth = intThisMonth - 1
	if intPrevMonth = 0 then
		intPrevMonth = 12
		intPrevYear = intThisYear - 1
	else
		intPrevYear = intThisYear	
	end if
		
	' Get the next month and year
	intNextMonth = intThisMonth + 1
	if intNextMonth > 12 then
		intNextMonth = 1
		intNextYear = intThisYear + 1
	else
		intNextYear = intThisYear
	end if

	' Get the last day of previous month. Using this, find the sunday of
	' last week of last month
	LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
	'The first day of the next month
	NextMonthDate = 1

	' Initialize the print day to 1  
	intPrintDay = 1

	' These dates are used in the SQL
	dFirstDay = intThisMonth & "/1/" & intThisYear & " 12:00:00 AM"
	dLastDay = intThisMonth & "/" & intLastDay & "/" & intThisYear & " 11:59:59 PM"

	'Open up a record set
	Query = "Select ID, Date, MemberID, Subject, StartDate, EndDate, Private, TotalRating, TimesRated FROM Calendar WHERE ( CustomerID = " & CustomerID & " AND " & _
			"( (StartDate >='" & dFirstDay & "' AND StartDate <= '" & dLastDay & "') " & _
			"OR " & _
			"(EndDate >='" & dFirstDay & "' AND EndDate <= '" & dLastDay & "') " & _
			"OR " & _
			"(StartDate < '" & dFirstDay & "' AND EndDate > '" & dLastDay & "' ) ) )"  & _
			"ORDER BY StartDate"
	'Create our connection object and open a connection to our database
	Set rsBDays = Server.CreateObject("ADODB.Recordset")
	rsBDays.CacheSize = 50
	Set rsEvents = Server.CreateObject("ADODB.Recordset")
	rsEvents.CacheSize = 50
	rsEvents.Open Query, Connect, adOpenStatic, adLockReadOnly

	Set ID = rsEvents("ID")
	Set ItemDate = rsEvents("Date")
	Set MemberID = rsEvents("MemberID")
	Set Subject = rsEvents("Subject")
	Set StartDate = rsEvents("StartDate")
	Set EndDate = rsEvents("EndDate")
	Set IsPrivate = rsEvents("Private")
	Set TotalRating = rsEvents("TotalRating")
	Set TimesRated = rsEvents("TimesRated")

	PrintTableHeader 100
	'------------------------End Code-----------------------------
	%>
		<tr>
			<td align="center" colspan="7" class="TDHeader">
				<table>
				<tr>
					<td class="TDHeader">
						<a HREF="calendar.asp?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><%=PrintTDLink( "Last Month" )%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</td>
					<td class="TDHeader">
					<%=strMonthName & " " & intThisYear %>
					</td>
					<td class="TDHeader">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a HREF="calendar.asp?month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>"><%=PrintTDLink( "Next Month")%></a>
					</td>
				</tr>
				</table>			
				<P>
				<FORM ACTION="calendar.asp" METHOD="POST">
					<select NAME="month" size="1">
	<%	
	'-----------------------Begin Code----------------------------	
						for intMonthNum = 1 to 12
							Response.Write("<option value='" & intMonthNum & "'")
							if intThisMonth = intMonthNum then Response.Write(" selected ")
							Response.Write(">" & MonthName(intMonthNum) & "</option>")
						next
	'------------------------End Code-----------------------------
	%>
					</select>
					<select NAME="year"  size="1">
	<%	
	'-----------------------Begin Code----------------------------	
						for intYear = 1998 to 2010
							Response.Write("<option value='" & intYear & "'")
							if intThisYear = intYear then Response.Write(" selected ")
							Response.Write(">" & intYear & "</option>")
						next
	'------------------------End Code-----------------------------
	%>
					</select>
					<input type="submit" CLASS="button" value="Go">
				</FORM>
				</p>
				Click on an event to get details.
			</td>
		</tr>
		<tr>
		<td class="<% PrintTDMain %>" align="center">Sun</td>
		<td class="<% PrintTDMain %>" align="center">Mon</td>
		<td class="<% PrintTDMain %>" align="center">Tue</td>
		<td class="<% PrintTDMain %>" align="center">Wed</td>
		<td class="<% PrintTDMain %>" align="center">Thu</td>
		<td class="<% PrintTDMain %>" align="center">Fri</td>
		<td class="<% PrintTDMain %>" align="center">Sat</td>	
		</tr>


	<%
	' Initialize the end of rows flag to false
	EndRows = False
				
	' loop until all the rows are exhausted
	do While EndRows = False
		' Start a table row
		Response.Write vbCrLf & "<tr>" & vbCrLf
		' This is the loop for the days in the week
		For intloopDay = cSUN To cSAT
			' if the first day is not sunday then print the last days of previous month in grayed font
			if intFirstWeekDay > cSUN then
				intLastMonth = intThisMonth - 1
				Write_TD_Faded
				LastMonthDate = LastMonthDate + 1
				intFirstWeekDay = intFirstWeekDay - 1

			' The month starts on a sunday or we are at a new row
			else
				' if the dates for the month are exhausted, start printing next month's dates as faded
				if intPrintDay > intLastDay then
					Write_TD_Faded
					NextMonthDate = NextMonthDate + 1
					EndRows = True 
				else
					' if last day of the month, flag the end of the row to print
					if intPrintDay = intLastDay then
						EndRows = True
					end if
					'Initialize the data to be printed
					strPrintData = ""
					'Set the current date
					dateCurrent = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)  
					'This will monitor the number of lines printed in a cell so we don't have uneven rows
					intLinesPrinted = 0
					'if we want to print out the member's birthdays, then check on it
					if CalendarShowBirthdays = 1 then
						Query = "Select ID FROM Members WHERE ( CustomerID = " & CustomerID & " AND " & _
								"( Day(Birthdate) ='" & intPrintDay & "' AND Month(Birthdate) = '" & intThisMonth & "') )"
						rsBDays.Open Query, Connect, adOpenStatic, adLockReadOnly
						do until rsBDays.EOF
							strPrintData = strPrintData & "<hr><a href='calendar_read.asp?ID=" & dateCurrent & "'>" & PrintTDLink( GetNickName(rsBDays("ID")) & "'s Birthday!" ) & "</a> <BR> "  & vbCrLf
							intLinesPrinted = intLinesPrinted + 1
							rsBDays.MoveNext
						loop
						rsBDays.Close
					end if
					rsEvents.Filter = "StartDate <= '" & dateCurrent & " 11:59:59 PM' AND EndDate >= '" & dateCurrent & " 12:00:00 AM'"
					if not rsEvents.EOF then
						do until rsEvents.EOF
							strPrivate = ""
							if IsPrivate = 1 and not blPrivateCategory and DisplayPrivacy then strPrivate = "Private, "

							strAuthor = ""
							if DisplayAuthor then strAuthor = "By: " & PrintTDLink(GetNickNameLink( MemberID )) & ", "

							strRating = ""
							if TimesRated > 0 and RateCalendar = 1 then strRating = "Rating: " & GetRating( TotalRating, TimesRated ) & ", "

							strDate = ""
							if DisplayDate then strDate = "Written: " & FormatDateTime(ItemDate, 2) & ", "

							strDetails = strAuthor & strPrivate & strDate & strRating
							if Len(strDetails) > 2 then strDetails = "<font size='-2'> ( " & Left( strDetails, (Len(strDetails) - 2) ) & " )</font> "


							strPrintData = strPrintData & "<hr><a href='calendar_read.asp?ID=" & dateCurrent & "'>" & PrintTDLink( Subject ) & "</a>" & strDetails & "<BR>"  & vbCrLf


							intLinesPrinted = intLinesPrinted + 1
							rsEvents.MoveNext
						loop
					end if

					if InStr( strPrintData, "<hr>" ) then
						strPrintData = Right( strPrintData, Len(strPrintData) - 4 )
					end if

					'Fill in the empty lines so each cell has at least 4 lines.  this makes the calendar vertically spaced right
					For FillLines = intLinesPrinted To 4
						strPrintData = strPrintData & " <BR> "  & vbCrLf
					Next
					
					' Print out today.
					strPrintData = "<BR>" & strPrintData
					'if the current date is today's actual date, then highlight it
					if dateCurrent = Date then 
	%>
						<td valign=top width="14%" class="TDMain2"><%=intPrintDay%><%=strPrintData%></td>
	<%
					else
	%>
						<td valign=top width="14%" class="<% PrintTDMain %>"><%=intPrintDay%><%=strPrintData%></td>
	<%
					end if
				end if 
						
				intPrintDay = intPrintDay + 1
			end if
				
		next
		Response.Write "</tr>" & vbCrLf
	loop 
	Set rsBDays = Nothing
	rsEvents.Close
	set rsEvents = Nothing
	%>

	</table>
<%

	'Give them the link to change the section's properties
	if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
		Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Calendar&Source=calendar.asp'>Change Section Options</a></p>"
	end if
end if

'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(Subject & Body & GetNickName(MemberID) )
End Function
%>

