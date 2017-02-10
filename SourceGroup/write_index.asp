<%
'-----------------------Create Files----------------------------
if strPath = "" then strPath = GetPath("")

'This allows us to include this in such files as nightmaint.asp, I'm dunk
if IsObject( FileSystem ) then
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	FileSystemObject = true
else
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	FileSystemObject = false
end if

Set IndexFile = FileSystem.CreateTextFile(strPath & "indextemp.inc")


'-----------------------Process shit----------------------------


if IsObject( rsPage ) then
	rsPageObject = true
else
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPageObject = false
end if

rsPage.CacheSize = 20

Query = "SELECT ListTypeNews, DisplayDateListNews, DisplayAuthorListNews FROM Look WHERE CustomerID = " & CustomerID & " ORDER BY Date DESC"
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

ListType = rsPage("ListTypeNews")
DisplayDate = CBool(rsPage("DisplayDateListNews"))
DisplayAuthor = CBool(rsPage("DisplayAuthorListNews"))
ItemNumber = 0	'This will be set by the PrintPagesHeader sub


if IsObject( FileSystem ) then
	blBulletImg = ImageExists("BulletImage", strBulletExt)
else
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	Set FileSystem = Nothing
end if

DblQuote = Chr(34)

rsPage.Close


Query = "SELECT Date, Body, MemberID  FROM News WHERE CustomerID = " & CustomerID & " ORDER BY Date DESC"
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect


'If they want to show birthdays, cool.  if not, force an empty recordset
if CalendarShowBirthdays = 1 and NewsShowEvents = 1 then
	Query = "Select ID FROM Members WHERE (CustomerID = " & CustomerID & " AND " & _
			"( Day(Birthdate) ='" & Day(Date) & "' AND Month(Birthdate) = '" & Month(Date) & "') )"
else
	Query = "Select ID FROM Members WHERE ID = 0"
end if

if IsObject( rsBDays ) then
	rsBDaysObject = true
else
	Set rsBDays = Server.CreateObject("ADODB.Recordset")
	rsBDaysObject = false
end if

rsBDays.CacheSize = 20
rsBDays.Open Query, Connect, adOpenStatic, adLockReadOnly

'If they want to show events, cool.  if not, force an empty recordset
if NewsShowEvents = 1 then
	dizay = FormatDateTime( Date, 2 )


	'we have to ignore the times, so we have to check the day, month and year individually.  only if sql had a formatdatetime..
	'if StartDate <= Today and EndDate >= Today

	Query = "SELECT Subject, MemberID FROM Calendar WHERE (CustomerID = " & CustomerID & " AND (MONTH(StartDate) <= " & Month(Date) & " AND MONTH(EndDate) >= " & Month(Date) & ") " & _
	" AND (YEAR(StartDate) <= " & Year(Date) & " AND YEAR(EndDate) >= " & Year(Date) & ") " & _
	" AND (DAY(StartDate) <= " & Day(Date) & " AND DAY(EndDate) >= " & Day(Date) & ") " & _
	") ORDER BY ID"
else
	Query = "Select NickName FROM Members WHERE ID = 0"
end if

if IsObject( rsCalendar ) then
	rsCalendarObject = true
else
	Set rsCalendar = Server.CreateObject("ADODB.Recordset")
	rsCalendarObject = false
end if
rsCalendar.CacheSize = 20
rsCalendar.Open Query, Connect, adOpenStatic, adLockReadOnly

if not (rsPage.EOF AND rsBDays.EOF AND rsCalendar.EOF) then

	IndexFile.WriteLine "<p align='" & HeadingAlignment & "'><span class='Heading'>" & NewsTitle & "</span>"

	IndexFile.WriteLine "<" & "% 'Link to add news"
	IndexFile.WriteLine "if IncludeAddButtons = 1 and LoggedAdmin() then " & _
						"Response.Write " & DblQuote & "<br><span class=LinkText><a href='admin_news_add.asp'>Add News</a></span>" & DblQuote
	IndexFile.WriteLine "%" & ">"
	IndexFile.WriteLine "</p>"

	'Print the list header
	if ListType = "Table" then
		IndexFile.WriteLine GetTableHeader(0)
		IndexFile.WriteLine "<tr>"
		if DisplayDate then IndexFile.WriteLine "<td class='TDHeader'>Date</td>"
		if DisplayAuthor then IndexFile.WriteLine "<td class='TDHeader'>Author</td>"
		IndexFile.WriteLine "<td class='TDHeader'>News</td></tr>"
	elseif ListType = "Bulleted" and not blBulletImg then
		IndexFile.WriteLine "<ul>"
	else
		IndexFile.WriteLine "<p>"
	end if

	if not rsBDays.EOF and CalendarShowBirthdays = 1 and NewsShowEvents = 1 then
		Set MembID = rsBDays("ID")
		do until rsBDays.EOF
			'Print the birthdays
			if ListType = "Table" then
					IndexFile.WriteLine "<tr>"
					if DisplayDate then IndexFile.WriteLine "<td class=" & GetTDMain & " align=center><font size=+1><b>Today</b></font></td>"
					if DisplayAuthor then IndexFile.WriteLine "<td class=" & GetTDMain & ">Everyone!</td>"
					IndexFile.WriteLine "<td class=" & GetTDMainSwitch & "><font size=+1><b>Happy birthday, " & GetNickName(MembID) & "</b></font></td></tr>"
			else
				strHeader = ""
				strFooter = "<br>"
				if ListType = "Bulleted" then
					if blBulletImg then
						strHeader = "<img src='images/BulletImage." & strBulletExt & "'>"
					else
						strHeader = "<li>"
						strFooter = "</li>"
					end if
				elseif ListType = "Numbered" then
						ItemNumber = ItemNumber + 1
						strHeader = ItemNumber & ".&nbsp;"
				end if
				IndexFile.WriteLine strHeader & "<font size=+1><b>Happy birthday, " & GetNickName(MembID) & "</b></font>&nbsp;&nbsp;&nbsp;&nbsp;"
				IndexFile.WriteLine  strFooter
			end if

			rsBDays.MoveNext
		loop
	end if






	if not rsCalendar.EOF and NewsShowEvents = 1 then
		Set Subject = rsCalendar("Subject")
		Set MemberID = rsCalendar("MemberID")
		do until rsCalendar.EOF
			'Print the birthdays
			if ListType = "Table" then
					IndexFile.WriteLine "<tr>"
					if DisplayDate then IndexFile.WriteLine "<td class=" & GetTDMain & " align=center><b>Today</b></td>"
					if DisplayAuthor then IndexFile.WriteLine "<td class=" & GetTDMainSwitch & ">" & PrintTDLink(GetNickNameLink(MemberID)) & "</td>"
					IndexFile.WriteLine "<td class=" & GetTDMainSwitch & "><b><a href='calendar_read.asp?ID= " & FormatDateTime(Date, 2) & "'>" & PrintTDLink(Subject) & "</a></b></td></tr>"
			else
				strHeader = ""
				strFooter = "<br>"
				if ListType = "Bulleted" then
					if blBulletImg then
						strHeader = "<img src='images/BulletImage." & strBulletExt & "'>"
					else
						strHeader = "<li>"
						strFooter = "</li>"
					end if
				elseif ListType = "Numbered" then
						ItemNumber = ItemNumber + 1
						strHeader = ItemNumber & ".&nbsp;"
				end if
				IndexFile.WriteLine strHeader
				if DisplayDate then IndexFile.WriteLine "Today - "
				IndexFile.WriteLine "<b><a href='calendar_read.asp?ID= " & FormatDateTime(Date, 2) & "'>" & Subject & "</a></b>&nbsp;&nbsp;&nbsp;&nbsp;"
				if DisplayAuthor then IndexFile.WriteLine "By: " & GetNickNameLink(MemberID) & "&nbsp;&nbsp;"
				IndexFile.WriteLine  strFooter
			end if

			rsCalendar.MoveNext
		loop
	end if

	if not rsPage.EOF then
		Set Body = rsPage("Body")
		Set MemberID = rsPage("MemberID")
		do until rsPage.EOF
			'Print the birthdays
			if ListType = "Table" then
					IndexFile.WriteLine "<tr>"
					if DisplayDate then IndexFile.WriteLine "<td class=" & GetTDMain & " align=center>" & FormatDateTime(rsPage("Date"), 2) & "</td>"
					if DisplayAuthor then IndexFile.WriteLine "<td class=" & GetTDMain & ">" & PrintTDLink(GetNickNameLink(MemberID)) & "</td>"
					IndexFile.WriteLine "<td class=" & GetTDMainSwitch & ">" & Body & "</td></tr>"
			else
				strHeader = ""
				strFooter = "<br>"
				if ListType = "Bulleted" then
					if blBulletImg then
						strHeader = "<img src='images/BulletImage." & strBulletExt & "'>"
					else
						strHeader = "<li>"
						strFooter = "</li>"
					end if
				elseif ListType = "Numbered" then
						ItemNumber = ItemNumber + 1
						strHeader = ItemNumber & ".&nbsp;"
				end if
				IndexFile.WriteLine strHeader
				if DisplayDate then IndexFile.WriteLine FormatDateTime(rsPage("Date"), 2) & " - "
				if DisplayAuthor then IndexFile.WriteLine "By: " & GetNickNameLink(MemberID) & "&nbsp;&nbsp;"
				IndexFile.WriteLine Body
				IndexFile.WriteLine  strFooter
			end if



			rsPage.MoveNext
		loop
		rsPage.MoveFirst
	end if

	'Print the list footer
	if ListType = "Table" then
		IndexFile.WriteLine "</table>"

	elseif ListType = "Bulleted" and not blBulletImg then
		IndexFile.WriteLine"</ul>"
	else
		IndexFile.WriteLine "</p>"
	end if

	DblQuote = Chr(34)

	if not rsPage.EOF then
		IndexFile.WriteLine "<" & "% 'Give them the link to change the section's properties"
		IndexFile.WriteLine "if LoggedAdmin() then " & _
							"Response.Write " & DblQuote & "<div align=right><a href='admin_news_modify.asp'>Modify News</a></div>" & DblQuote
		IndexFile.WriteLine "%" & ">"
	end if

	IndexFile.WriteLine "<" & "% 'Give them the link to change the section's properties"
	IndexFile.WriteLine "if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then " & _
						"Response.Write " & DblQuote & "<div align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=News'>Change " & NewsTitle & " Options</a></div>" & DblQuote
	IndexFile.WriteLine "%" & ">"


end if



rsPage.Close


'Open up the item
Query = "SELECT ID, Title, Body FROM InfoPages WHERE CustomerID = " & CustomerID & " AND Title = 'Home Page'"
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

if not rsPage.EOF then
	if IsObject(rsPage("Body")) then
		if not IsNull(rsPage("Body")) then
			if rsPage("Body") <> "" then
				IndexFile.WriteLine "<br>" & rsPage("Body") & "<br>"
					IndexFile.WriteLine "<" & "% 'Give them the link to change the section's properties"
					IndexFile.WriteLine "if LoggedAdmin() then " & _
					"Response.Write " & DblQuote & "<div align=right><a href='members_pages_modify.asp?Submit=Edit&ID=" & rsPage("ID") & "'>Edit Home Page</a></div>" & DblQuote
					IndexFile.WriteLine "%" & ">"
			end if
		end if
	end if
end if

rsPage.Close
rsCalendar.Close
rsBDays.Close


IndexFile.Close
Set IndexFile = Nothing

FileSystem.CopyFile strPath & "indextemp.inc", strPath & "index.inc"

'If we created the objects, kill them. If not, leave em alone
if not FileSystemObject then Set FileSystem = Nothing

if not rsPageObject then Set rsPage = Nothing


if not rsBDaysObject then Set rsBDays = Nothing

if not rsCalendarObject then Set rsCalendar = Nothing


%>
