
<%
'Version 1 - 12/9/2000 - Stephen Potter
'-----------------------Begin Code----------------------------
if IncludeAdditions = 1 then
'------------------------End Code-----------------------------
%>
	<script language="JavaScript"><!--
	function ValidateInteger(strCheck) {
		if (!strCheck) return false;
		if (strCheck.search(/^[0-9]*$/) != -1)
			 return true;
		 else
			 return false;
	}


	function submit_page(form) {
		if (ValidateInteger(form.DaysOld.value)){
			return true;
		}
		else{
			alert ('Sorry, but you must enter a valid number of days.');
			return false;
		}   
	}
	//-->
	</SCRIPT>


	<p class="Heading" align="<%=HeadingAlignment%>"><%=AdditionsTitle%></p>
	<form METHOD="post" ACTION="additions.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	View additions in the last <input type="text" size="2" name="DaysOld"> days. <input type="submit" value="Go">
	</form>
<%
'-----------------------Begin Code----------------------------
buffer = false

Function GetDisplay( blDisplay )
	blDisplay = CBool(blDisplay)
	if blDisplay then
		GetDisplay = ""
	else
		GetDisplay = " style=" & chr(34) & "display: none;" & chr(34)
	end if
End Function

if Request("DaysOld") <> "" then
	intDaysOld = CInt(Request("DaysOld"))
else
	intDaysOld = AdditionsDaysOld
end if

boolEmpty = true

CutoffDate = DateAdd("d", (-1*intDaysOld ), Date)

Set rsLatest = Server.CreateObject("ADODB.Recordset")
rsLatest.CacheSize = 60
Set rsCat = Server.CreateObject("ADODB.Recordset")


if IncludeAdditionsMembers = 1 or IncludeAdditionsMembers = "" then


	Query = "SELECT ID FROM Members WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then 
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('membersChild'); return false">
			New Members <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="membersChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader">New Members <%=strNum%></td></tr>
		<% end if

		boolEmpty = false
		do until rsLatest.EOF
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<b><%=GetNickNameLink( ID )%></b> has been added</a>.  
			</td></tr>
<%
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close

end if

if IncludeAdditionsInfoPages = 1 or IncludeAdditionsInfoPages = "" then
	blAuthors = IncludeAdditionsInfoPagesAuthor = 1 or IncludeAdditionsInfoPagesAuthor = ""

	Query = "SELECT ID, Title, MemberID FROM InfoPages WHERE ( Title <> 'Home Page' AND CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then 
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set PageTitle = rsLatest("Title")
		Set MemberID = rsLatest("MemberID")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('InfoPagesChild'); return false">
			Information Pages <%=strNum%></a>
			</td></tr></table>
	
			<span <%=GetDisplay(0)%> ID="InfoPagesChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader">Text Pages <%=strNum%></td></tr>
		<% end if

		boolEmpty = false
		do until rsLatest.EOF
%>
			<tr><td class="<% PrintTDMainSwitch %>">

			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added an information page called  
			<% end if %>
			<a href="<%=NonSecurePath%>pages_read.asp?ID=<%=ID%>"><b><%=PrintTDLink(PrintStart( PageTitle ))%></b></a>.  
			</td></tr>
<%
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new announcements------------------------
if CBool( IncludeAnnouncements and ( IncludeAdditionsAnnouncements = 1 or IncludeAdditionsAnnouncements = "" ) ) then
	blAuthors = IncludeAdditionsAnnouncementsAuthor = 1 or IncludeAdditionsAnnouncementsAuthor = ""
	Query = "SELECT ID, MemberID, Subject FROM Announcements WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then 
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('announcementsChild'); return false">
			<%=AnnouncementsTitle%> <%=strNum%></a>

			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="announcementsChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=AnnouncementsTitle%> <%=strNum%></td></tr>
		<% end if

		boolEmpty = false
		do until rsLatest.EOF
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added an announcement called  
			<% end if %>
			<a href="<%=NonSecurePath%>announcements_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


if CBool( IncludeMeetings and ( IncludeAdditionsMeetings = 1 or IncludeAdditionsMeetings = "" ) ) then
	blAuthors = IncludeAdditionsMeetingsAuthor = 1 or IncludeAdditionsMeetingsAuthor = ""

	Query = "SELECT ID, MemberID, Subject FROM Meetings WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then 
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('meetingsChild'); return false">
			<%=MeetingsTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="meetingsChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=MeetingsTitle%> <%=strNum%></td></tr>
		<% end if

		boolEmpty = false
		do until rsLatest.EOF
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a meeting called  
			<% end if %>

			<a href="<%=NonSecurePath%>meetings_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new stories------------------------
if CBool( IncludeStories ) and ( IncludeAdditionsStories = 1 or IncludeAdditionsStories = "" ) then
	blAuthors = IncludeAdditionsStoriesAuthor = 1 or IncludeAdditionsStoriesAuthor = ""

	Query = "SELECT ID, MemberID, Subject FROM Stories WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('StoriesChild'); return false">
			<%=StoriesTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="StoriesChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=StoriesTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a story called  
			<% end if %>

			<a href="<%=NonSecurePath%>stories_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new calendar events------------------------
if CBool( IncludeCalendar ) and ( IncludeAdditionsCalendar = 1 or IncludeAdditionsCalendar = "" ) then
	blAuthors = IncludeAdditionsCalendarAuthor = 1 or IncludeAdditionsCalendarAuthor = ""

	Query = "SELECT ID, MemberID, Subject, StartDate, EndDate FROM Calendar WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		Set StartDate = rsLatest("StartDate")
		Set EndDate = rsLatest("EndDate")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('CalendarChild'); return false">
			<%=CalendarTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="CalendarChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=CalendarTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			if StartDate = EndDate then
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added an event for <%=FormatDateTime(StartDate, 2)%> called 
			<% else %>
				An event for <%=FormatDateTime(StartDate, 2)%> called 
			<% end if %>

			<a href="<%=NonSecurePath%>calendar_event_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added an event for <%=FormatDateTime(StartDate, 2)%> - <%=FormatDateTime(EndDate, 2)%> called 
			<% else %>
				An event for <%=FormatDateTime(StartDate, 2)%> - <%=FormatDateTime(EndDate, 2)%> called 
			<% end if %>
			<a href="<%=NonSecurePath%>calendar_event_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			end if
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new links------------------------
if CBool( IncludeLinks ) and ( IncludeAdditionsLinks = 1 or IncludeAdditionsLinks = "" ) then
	blAuthors = IncludeAdditionsLinksAuthor = 1 or IncludeAdditionsLinksAuthor = ""

	Query = "SELECT ID, MemberID, URL, Name FROM Links WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set URL = rsLatest("URL")
		Set Name = rsLatest("Name")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('LinksChild'); return false">
			<%=LinksTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="LinksChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=LinksTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			strName = Name
			if strName = "" then strName = URL
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a link called  
			<% end if %>
			<a href="<%=NonSecurePath%>links_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( strName ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new stories------------------------
if CBool( IncludeQuotes ) and ( IncludeAdditionsQuotes = 1 or IncludeAdditionsQuotes = "" ) then
	blAuthors = IncludeAdditionsQuotesAuthor = 1 or IncludeAdditionsQuotesAuthor = ""

	Query = "SELECT ID, MemberID, Author, Quote FROM Quotes WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Quote = rsLatest("Quote")
		Set Author = rsLatest("Author")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('QuotesChild'); return false">
			<%=QuotesTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="QuotesChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=QuotesTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a quote by 
			<% end if %>

			<%=Author%> - <a href="<%=NonSecurePath%>quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink( "&quot;<b>"&PrintStart(Quote)&"</b>&quot;")%></a>
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new forum messages------------------------
if CBool( IncludeForum ) and ( IncludeAdditionsForumMessages = 1 or IncludeAdditionsForumMessages = "" ) then
	blAuthors = IncludeAdditionsForumAuthor = 1 or IncludeAdditionsForumAuthor = ""

	Query = "SELECT ID, MemberID, Subject, Author, Email, CategoryID FROM ForumMessages WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "' AND CategoryID > 0) ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		Set Author = rsLatest("Author")
		Set Email = rsLatest("Email")
		Set CategoryID = rsLatest("CategoryID")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('ForumChild'); return false">
			<%=ForumTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="ForumChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=ForumTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			if MemberID > 0 then
				strAuthor = PrintTDLink( GetNickNameLink(MemberID) )
			elseif InStr(Email, "@" ) then
				strAuthor = "<a href='mailto:" & Email & "'>" & PrintTDLink( Author ) & "</a>"
			else
				strAuthor = Author
			end if

			Query = "SELECT Name FROM ForumCategories WHERE ID = " & CategoryID
			rsCat.Open Query, Connect, adOpenStatic, adLockReadOnly
			strCat = "unknown"
			if not rsCat.EOF then strCat = rsCat("Name")
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=strAuthor%> has added a message to <%=strCat%> called  
			<% end if %>
			<a href="<%=NonSecurePath%>forum_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsCat.Close
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new stories------------------------
if CBool( IncludeVoting ) and ( IncludeAdditionsVotingPolls = 1 or IncludeAdditionsVotingPolls = "" ) then
	blAuthors = IncludeAdditionsVotingAuthor = 1 or IncludeAdditionsVotingAuthor = ""

	Query = "SELECT ID, MemberID, Subject FROM VotingPolls WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('VotingChild'); return false">
			<%=VotingTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="VotingChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=VotingTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a voting poll called  
			<% end if %>
			<a href="<%=NonSecurePath%>voting_cast.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new stories------------------------
if CBool( IncludeQuizzes ) and ( IncludeAdditionsQuizzes = 1 or IncludeAdditionsQuizzes = "" ) then
	blAuthors = IncludeAdditionsQuizzesAuthor = 1 or IncludeAdditionsQuizzesAuthor = ""

	Query = "SELECT ID, MemberID, Subject FROM Quizzes WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('QuizzesChild'); return false">
			<%=QuizzesTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="QuizzesChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=QuizzesTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a quiz called  
			<% end if %>
			<a href="<%=NonSecurePath%>quizzes_take.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Subject ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


'---------------Get new Guestbook------------------------
if CBool( IncludeGuestbook ) and ( IncludeAdditionsGuestbook = 1 or IncludeAdditionsGuestbook = "" ) then
	blAuthors = IncludeAdditionsGuestbookAuthor = 1 or IncludeAdditionsGuestbookAuthor = ""

	Query = "SELECT ID, Body, Email, Author FROM Guestbook WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set Body = rsLatest("Body")
		Set Email = rsLatest("Email")
		Set Author = rsLatest("Author")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('GuestbookChild'); return false">
			<%=GuestbookTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="GuestbookChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=GuestbookTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
				if InStr( Email, "@" ) then
					strAuthor = "<a href='mailto:" & Email & "'>" & PrintTDLink( Author ) & "</a>"
				else
					strAuthor = Author
				end if
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=strAuthor%> has added an entry:  
			<% end if %>
			<a href="<%=NonSecurePath%>guestbook_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( Body ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if



'---------------Get new stories------------------------
if CBool( IncludeMedia ) and ( IncludeAdditionsMedia = 1 or IncludeAdditionsMedia = "" ) then
	blAuthors = IncludeAdditionsMediaAuthor = 1 or IncludeAdditionsMediaAuthor = ""

	Query = "SELECT ID, MemberID, Description, FileName, CategoryID FROM Media WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "' AND CategoryID > 0) ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Description = rsLatest("Description")
		Set FileName = rsLatest("FileName")
		Set CategoryID = rsLatest("CategoryID")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('MediaChild'); return false">
			<%=MediaTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="MediaChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=MediaTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			strName = Description
			if strName = "" then strName = FileName
			Query = "SELECT Name FROM MediaCategories WHERE ID = " & CategoryID
			rsCat.Open Query, Connect, adOpenStatic, adLockReadOnly
			strCat = "unknown"
			if not rsCat.EOF then strCat = rsCat("Name")
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
			<% if blAuthors then %>
				<%=PrintTDLink( GetNickNameLink(MemberID) )%> has added a file to <%=strCat%> called  
			<% else %>
				A file to <%=strCat%> called  
			<% end if %>

			<a href="<%=NonSecurePath%>media_read.asp?ID=<%=ID%>"><b><%=PrintTDLink( PrintStart( strName ) )%></b></a>.  
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsCat.Close
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if


PrintCustom


'---------------Get new stories------------------------
if CBool( IncludePhotos ) and ( IncludeAdditionsPhotos = 1 or IncludeAdditionsPhotos = "" ) then
	blAuthors = IncludeAdditionsPhotosAuthor = 1 or IncludeAdditionsPhotosAuthor = ""

	Query = "SELECT ID, MemberID, Name, Thumbnail, ThumbnailExt, CategoryID FROM Photos WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') AND CategoryID > 0 ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set MemberID = rsLatest("MemberID")
		Set Name = rsLatest("Name")
		Set Thumbnail = rsLatest("Thumbnail")
		Set ThumbnailExt = rsLatest("ThumbnailExt")
		Set CategoryID = rsLatest("CategoryID")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('PhotosChild'); return false">
			<%=PhotosTitle%> <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="PhotosChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader"><%=PhotosTitle%> <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			Query = "SELECT Name FROM PhotoCategories WHERE ID = " & CategoryID
			rsCat.Open Query, Connect, adOpenStatic, adLockReadOnly
			strCat = "unknown"
			if not rsCat.EOF then strCat = rsCat("Name")
'------------------------End Code-----------------------------
%>
			<tr><td class="<% PrintTDMainSwitch %>">
<%
			if Thumbnail = 1 and not PhotosPrivateCategory( CategoryID ) then
%>
				<% if blAuthors then %>
					<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added this photo to <%=strCat%>:   
				<% else %>
					A photo to <%=strCat%>:   
				<% end if %>


				<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=ID%>"><img src="photos/<%=ID%>t.<%=ThumbnailExt%>" border=0 alt="<%=Name%>"></a>
<%
			elseif Thumbnail = 1 then
%>
				<% if blAuthors then %>
					<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a 
				<% end if %>

				<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=ID%>"><%=PrintTDLink( "<b>private photo</b></a> to " & strCat)%>.  
<%
			else
%>
				<% if blAuthors then %>
					<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a photo to <%=strCat%> called  
				<% else %>
					A photo to <%=strCat%> called  
				<% end if %>
				<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=ID%>"><%=PrintTDLink( "<b>" & PrintStart(Name) & "</b>" )%></a>.  
<%
			end if
%>
			</td></tr>
<%
'-----------------------Begin Code----------------------------
			rsCat.Close
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close

	if ( IncludeAdditionsPhotoCaptions = 1 or IncludeAdditionsPhotoCaptions = "" ) then
		blAuthors = IncludeAdditionsPhotoCaptionsAuthor = 1 or IncludeAdditionsPhotoCaptionsAuthor = ""

		Query = "SELECT ID, MemberID, PhotoID, Caption FROM PhotoCaptions WHERE ( CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsLatest.EOF then
			intNum = rsLatest.RecordCount
			if intNum = 1 then
				strNum = " <font size=-2>(1 Addition)</font>"
			else
				strNum = " <font size=-2>(" & intNum & " Additions)</font>"
			end if
			Set ID = rsLatest("ID")
			Set MemberID = rsLatest("MemberID")
			Set PhotoID = rsLatest("PhotoID")
			Set Caption = rsLatest("Caption")
			PrintTableHeader 100
			%>
			<br>
			<% if intNum > 5 then %>
				<tr><td class="TDHeader">
				<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('PhotoCaptionsChild'); return false">
				<%=PhotoCaptionsTitle%> <%=strNum%></a>
				</td></tr></table>
				<span <%=GetDisplay(0)%> ID="PhotoCaptionsChild">
			<% 	PrintTableHeader 100
			else %>
				<tr><td class="TDHeader"><%=PhotoCaptionsTitle%> <%=strNum%></td></tr>
			<% end if
			boolEmpty = false
			do until rsLatest.EOF
				Query = "SELECT Name, ID, Thumbnail, ThumbnailExt, CategoryID FROM Photos WHERE ID = " & PhotoID
				rsCat.Open Query, Connect, adOpenStatic, adLockReadOnly
	'------------------------End Code-----------------------------
	%>
				<tr><td class="<% PrintTDMainSwitch %>">
	<%
				if rsCat("Thumbnail") = 1 and not PhotosPrivateCategory( rsCat("CategoryID") ) then
	%>
					<% if blAuthors then %>
						<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added added a caption to this photo:   
					<% else %>
						A caption to this photo:   
					<% end if %>

					<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=PhotoID%>"><img src="photos/<%=rsCat("ID")%>t.<%=rsCat("ThumbnailExt")%>" border=0 alt="<%=rsCat("Name")%>"></a>
	<%
				elseif rsCat("Thumbnail") = 1 then
	%>
					<% if blAuthors then %>
						<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a caption: 
					<% else %>
						A caption:
					<% end if %>

					<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=PhotoID%>"><b><%=PrintTDLink( PrintStart( Caption ) )%></b></a>.  
	<%
				else
	%>
					<% if blAuthors then %>
						<%=PrintTDLink( GetNickNameLink(MemberID ) )%> has added a caption to a photo called  
					<% else %>
						A caption to a photo called  
					<% end if %>
					<a href="<%=NonSecurePath%>photos_view.asp?ID=<%=PhotoID%>"><b><%=PrintTDLink( PrintStart( rsCat("Name") ) )%></b></a>.  
	<%
				end if
	%>
				</td></tr>
	<%
	'-----------------------Begin Code----------------------------
				rsCat.Close
				rsLatest.MoveNext
			loop
			Response.Write "</table>"
			if intNum > 5 then Response.Write "</span>"
		end if
		rsLatest.Close
	end if
end if

'---------------Get new stories------------------------
if ( IncludeAdditionsReviews = 1 or IncludeAdditionsReviews = "" ) then
	blAuthors = IncludeAdditionsReviewsAuthor = 1 or IncludeAdditionsReviewsAuthor = ""

	Query = "SELECT ID, TargetTable, TargetID, MemberID, Subject, Author, Email FROM Reviews WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	rsLatest.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsLatest.EOF then
		intNum = rsLatest.RecordCount
		if intNum = 1 then
			strNum = " <font size=-2>(1 Addition)</font>"
		else
			strNum = " <font size=-2>(" & intNum & " Additions)</font>"
		end if
		Set ID = rsLatest("ID")
		Set TargetTable = rsLatest("TargetTable")
		Set TargetID = rsLatest("TargetID")
		Set MemberID = rsLatest("MemberID")
		Set Subject = rsLatest("Subject")
		Set Author = rsLatest("Author")
		Set Email = rsLatest("Email")
		PrintTableHeader 100
		%>
		<br>
		<% if intNum > 5 then %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('ReviewsChild'); return false">
			Reviews <%=strNum%></a>
			</td></tr></table>
			<span <%=GetDisplay(0)%> ID="ReviewsChild">
		<% 	PrintTableHeader 100
		else %>
			<tr><td class="TDHeader">Reviews <%=strNum%></td></tr>
		<% end if
		boolEmpty = false
		do until rsLatest.EOF
			if MemberID > 0 then
				strAuthor = PrintTDLink( GetNickNameLink(MemberID ) )
			elseif InStr(Email, "@" ) then
				strAuthor = "<a href='mailto:" & Email & "'>" & PrintTDLink( Author ) & "</a>"
			else
				strAuthor = Author
			end if

			if blAuthors then
				strAuthor = strAuthor & " has reviewed "
			else
				strAuthor = "A review for "

			end if

			if TargetTable = "Announcements" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"an announcement called <a href='" & NonSecurePath & "announcements_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Calendar" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a calendar event called <a href='" & NonSecurePath & "calendar_event_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Meetings" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a meeting called <a href='" & NonSecurePath & "meetings_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Guestbook" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a guestbook entry called <a href='" & NonSecurePath & "guestbook_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Links" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a link called <a href='" & NonSecurePath & "links_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Media" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a file called <a href='" & NonSecurePath & "media_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Members" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				GetNickNameLink(TargetID) & " called <a href='" & NonSecurePath & "member.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "PhotoCaptions" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a caption called <a href='" & NonSecurePath & "photocaptions_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Quizzes" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a quiz called <a href='" & NonSecurePath & "quizzes_rate.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Quotes" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a quote called <a href='" & NonSecurePath & "quotes_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "Stories" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a story called <a href='" & NonSecurePath & "stories_read.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			elseif TargetTable = "VotingPolls" then
				Response.Write "<tr><td class=" & GetTDMainSwitch & ">" & _
				strAuthor & _
				"a voting poll called <a href='" & NonSecurePath & "voting_results.asp?ID=" & TargetID & "'>" & _
				"<b>" & PrintTDLink( PrintStart( Subject ) ) & "</b></a>.</td></tr>"
			end if
			PrintCustomReviews
			rsLatest.MoveNext
		loop
		Response.Write "</table>"
		if intNum > 5 then Response.Write "</span>"
	end if
	rsLatest.Close
end if

if boolEmpty = true then
%>
	<p class="<% PrintTDMain %>">None right now.</p>
<%
end if



'Give them the link to change the section's properties
if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
	Response.Write "<br><br><p align=right><a href='admin_additions_configure.asp?Source=announcements.asp'>Configure Additions</a></p>"
end if

Set rsLatest = Nothing
Set rsCat = Nothing

'This is the if statement that includes the additions section or not.  Not going to indent everything though
end if


Function PhotosPrivateCategory( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PhotosPrivateCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Private", adInteger, adParamOutput )

		.Parameters("@ItemID") = intCategoryID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Private")
	End With
	Set cmdTemp = Nothing

	PhotosPrivateCategory = CBool(blExists)
End Function

%>