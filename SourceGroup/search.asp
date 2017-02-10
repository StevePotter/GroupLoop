<p class="Heading" align="<%=HeadingAlignment%>">Search</p>

<%
Server.ScriptTimeout = 5400

strSearch = Request("Keywords")
intSearchID = Request("SearchID")

'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if strSearch <> "" then
	Set cmdSearch = Server.CreateObject("ADODB.Command")
	With cmdSearch
		.ActiveConnection = Connect
		.CommandText = "AddSearchRow"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@SearchIDOut", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@SearchID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Score", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Author", adVarWChar, adParamInput, 200 )
		.Parameters.Append .CreateParameter ("@Link", adVarWChar, adParamInput, 200 )
		.Parameters.Append .CreateParameter ("@Type", adVarWChar, adParamInput, 200 )
	End With

	IncrementStat("Searches")

	strSearch = UCASE(strSearch)
	strTables = Request("Tables")
	intSearchID = 0

	Set rsList = Server.CreateObject("ADODB.Recordset")
	rsList.CacheSize = 1000

	'Get the keywords and break them up
	Dim strWords(100)
	intWordCounter = 1
	'The words are separated by commas or spaces.
	for i = 1 to len(strSearch)
	   if mid(strSearch, i, 1) = " " Or mid(strSearch, i, 1) = "," then 
		 intWordCounter = intWordCounter + 1
	   else
		 strWords(intWordCounter) = strWords(intWordCounter) & mid(strSearch, i, 1)
	   end if
	next


	if (strTables = "All" or InStr(strTables,"Announcements")) and CBool( IncludeAnnouncements ) then
		Query = "SELECT ID, MemberID, Subject, Body FROM Announcements WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = AnnouncementsTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='announcements_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = AnnouncementsTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='announcements_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Newsletters")) and CBool( IncludeNewsletter )  then
		Query = "SELECT ID, MemberID, Subject, Body FROM Newsletters WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = NewsletterTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='newsletter_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = NewsletterTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='newsletter_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"InfoPages")) then
		Query = "SELECT ID, MemberID, Title, Body FROM InfoPages WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set PageTitle = rsList("Title")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(PageTitle & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = "Text Pages"
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='pages_read.asp?ID=" & ID & "'>" & PrintStart(PageTitle) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = "Text Pages"
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='pages_read.asp?ID=" & ID & "'>" & PrintStart(PageTitle) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Stories")) and CBool( IncludeStories )  then
		Query = "SELECT ID, MemberID, Subject, Body FROM Stories WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = StoriesTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='stories_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = StoriesTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='stories_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Calendar")) and CBool( IncludeCalendar )  then
		Query = "SELECT ID, MemberID, Subject, Body FROM Calendar WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = CalendarTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='calendar_event_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = CalendarTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='calendar_event_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Forum")) and CBool( IncludeForum )  then
		Query = "SELECT ID, MemberID, Author, Subject, Body FROM ForumMessages WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Author = rsList("Author")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject & Body )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = ForumTitle
					if MemberID > 0 then
						strAuthor = GetNickNameLink( MemberID )
					else
						strAuthor = Author
					end if
					strLink = "<a href='forum_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = ForumTitle
						if MemberID > 0 then
							strAuthor = GetNickNameLink( MemberID )
						else
							strAuthor = Author
						end if
						strLink = "<a href='forum_read.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if



	if (strTables = "All" or InStr(strTables,"Links")) and CBool( IncludeLinks )  then
		Query = "SELECT ID, MemberID, URL, Name, Description FROM Links WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set URL = rsList("URL")
			Set Name = rsList("Name")
			Set Description = rsList("Description")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(URL & Name & Description )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = LinksTitle
					strAuthor = GetNickNameLink( MemberID )
					strName = Name
					if strName = "" then strName = URL
					strLink = "<a href='links_read.asp?ID=" & ID & "'>" & PrintStart(strName) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = LinksTitle
						strAuthor = GetNickNameLink( MemberID )
						strName = Name
						if strName = "" then strName = URL
						strLink = "<a href='links_read.asp?ID=" & ID & "'>" & PrintStart(strName) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Quotes")) and CBool( IncludeQuotes )  then
		Query = "SELECT ID, MemberID, Author, Quote, Description FROM Quotes WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Author = rsList("Author")
			Set Quote = rsList("Quote")
			Set Description = rsList("Description")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Author & Quote & Description )
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = QuotesTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='quotes_read.asp?ID=" & ID & "'>&quot;" & PrintStart(Quote) & "&quot;</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = QuotesTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='quotes_read.asp?ID=" & ID & "'>" & PrintStart(strName) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Photos")) and CBool( IncludePhotos )  then
		Query = "SELECT ID, MemberID, Name, Thumbnail, ThumbnailExt FROM Photos WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Name = rsList("Name")
			Set Thumbnail = rsList("Thumbnail")
			Set ThumbnailExt = rsList("ThumbnailExt")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Name)
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = PhotosTitle
					strAuthor = GetNickNameLink( MemberID )
					if Thumbnail = 1 then
						strLink = "<a href='photos_view.asp?ID=" & ID & "'><img src='photos/" & ID & "t." & ThumbnailExt & "' border=0 alt='" & Name & "'></a>"
					else
						strLink = "<a href='photos_view.asp?ID=" & ID & "'>" & PrintStart(Name) & "</a>"
					end if
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = PhotosTitle
						strAuthor = GetNickNameLink( MemberID )
						if Thumbnail = 1 then
							strLink = "<a href='photos_view.asp?ID=" & ID & "'><img src='photos/" & ID & "t." & ThumbnailExt & "' border=0 alt='" & Name & "'></a>"
						else
							strLink = "<a href='photos_view.asp?ID=" & ID & "'>" & PrintStart(Name) & "</a>"
						end if
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if

	if (strTables = "All" or InStr(strTables,"Photos")) and CBool( IncludePhotos ) and CBool( IncludePhotoCaptions )  then
		Query = "SELECT ID, MemberID, Caption FROM PhotoCaptions WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Caption = rsList("Caption")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Caption)
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = PhotoCaptionsTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='photocaptions_read.asp?ID=" & ID & "'>" & PrintStart(Caption) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = PhotoCaptionsTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='photocaptions_read.asp?ID=" & ID & "'>" & PrintStart(Caption) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Quizzes")) and CBool( IncludeQuizzes )  then
		Query = "SELECT ID, MemberID, Subject FROM Quizzes WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject)
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = QuizzesTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='quizzes_take.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = QuizzesTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='quizzes_take.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if


	if (strTables = "All" or InStr(strTables,"Voting")) and CBool( IncludeVoting )  then
		Query = "SELECT ID, MemberID, Subject FROM VotingPolls WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Subject)
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = VotingTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='voting_cast.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = VotingTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='voting_cast.asp?ID=" & ID & "'>" & PrintStart(Subject) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if

	if (strTables = "All" or InStr(strTables,"Store")) and CBool( IncludeStore )  then
		Query = "SELECT ID, MemberID, Name, Description, Note FROM StoreGroups WHERE (CustomerID = " & CustomerID & ")"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsList.EOF then
			Set ID = rsList("ID")
			Set MemberID = rsList("MemberID")
			Set Name = rsList("Name")
			Set Description = rsList("Description")
			Set Note = rsList("Note")

			'Now go through, looking for matches
			do until rsList.EOF
				strDescription = UCASE(Name&Description&Note)
				'check for the whole search string in the description
				if InStr(strDescription,strSearch) then
					strType = StoreTitle
					strAuthor = GetNickNameLink( MemberID )
					strLink = "<a href='store_view.asp?ID=" & ID & "'>" & PrintStart(Name) & "</a>"
					intSearchID = AddMatch( intSearchID, 100, CustomerID, ItemDate, strAuthor, strLink, strType )
				'Now check word for word
				else
					intNumMatches = 0
					for i = 1 to intWordCounter
						if InStr(strDescription,strWords(i)) then 
							intNumMatches = intNumMatches + 1
						end if
					next
					'We got a match, so add it
					if intNumMatches > 0 then
						intScore = Round( 100 * intNumMatches / intWordCounter)
						if intScore = 100 then intScore = 99
						strType = StoreTitle
						strAuthor = GetNickNameLink( MemberID )
						strLink = "<a href='store_view.asp?ID=" & ID & "'>" & PrintStart(Name) & "</a>"
						intSearchID = AddMatch( intSearchID, intScore, CustomerID, ItemDate, strAuthor, strLink, strType )
					end if
				end if
				rsList.MoveNext
			loop
		end if
		rsList.Close
	end if

	Set cmdSearch = Nothing

	Set rsList = Nothing
end if


'They haven't entered anything yet
if Request("Keywords") = "" and Request("SearchID") = "" then
%>
	<p>You may search all sections or just certain ones.  To select more than one section, just hold down the Control key while selecting them.  Please keep in mind that a full search may take some, so please click 'Go' once and be patient.</p>
	<form METHOD="POST" ACTION="search.asp">
	<table cellspacing=2 cellpadding=2>
	<tr>
	<td valign=top>
		Search 
	</td>
	<td valign=top>
		<select name="Tables" size="3" multiple>
			<option value="All" selected>All Sections</option>
<%
	Response.Write	"<option value='InfoPages'>Text Pages</option>"
	if CBool( IncludeAnnouncements ) then Response.Write	"<option value='Announcements'>" & AnnouncementsTitle & "</option>"
	if CBool( IncludeNewsletter ) then Response.Write	"<option value='Newsletters'>" & NewsletterTitle & "</option>"
	if CBool( IncludeStories ) then Response.Write	"<option value='Stories'>" & StoriesTitle & "</option>"
	if CBool( IncludeCalendar ) then Response.Write	"<option value='Calendar'>" & CalendarTitle & "</option>"
	if CBool( IncludeLinks ) then Response.Write	"<option value='Links'>" & LinksTitle & "</option>"
	if CBool( IncludeQuotes ) then Response.Write	"<option value='Quotes'>" & QuotesTitle & "</option>"
	if CBool( IncludeForum ) then Response.Write	"<option value='Forum'>" & ForumTitle & "</option>"
	if CBool( IncludePhotos ) then Response.Write	"<option value='Photos'>" & PhotosTitle & "</option>"
	if CBool( IncludeVoting ) then Response.Write	"<option value='Voting'>" & VotingTitle & "</option>"
	if CBool( IncludeStore ) and CBool( AllowStore ) then Response.Write	"<option value='Store'>" & StoreTitle & "</option>"
	if CBool( IncludeQuizzes ) then Response.Write	"<option value='Quizzes'>" & QuizzesTitle & "</option>"
%>
		</select>
	</td>
	<td valign=top>
		For 
	</td>
	<td valign=top>
		<input type="text" name="Keywords" size="25"> <input type="submit" name="Submit" value="Go">
	</td>
	</table>
	</form>
<%
else
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>, but search came up empty.<br>
			<a href="search.asp">Click here</a> to try again.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			<a href="search.asp">Click here</a> to try again.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT Score, Author, Type, Link FROM SiteSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC, Type"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
		<form METHOD="POST" ACTION="search.asp">

		<input type="hidden" name="Tables" value="All">
		<p>Your search resulted in <b><%=rsPage.RecordCount%></b> matches.<br>
		Search again for <input type="text" name="Keywords" size="25"> <input type="submit" name="Submit" value="Go"></p>
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">

<%
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Match Score</td>
			<td class="TDHeader">Author</td>
			<td class="TDHeader">Section</td>
			<td class="TDHeader">Item</td>
		</tr>
<%
		Set Score = rsPage("Score")
		Set Author = rsPage("Author")
		Set ItemType = rsPage("Type")
		Set Link = rsPage("Link")
		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMain %>" align="center"><%=Score%></td>
					<td class="<% PrintTDMain %>"><%=Author%></td>
					<td class="<% PrintTDMain %>"><%=ItemType%></td>
					<td class="<% PrintTDMainSwitch %>"><%=Link%></td>
				</tr>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
		set rsPage = Nothing
	end if
end if


Function AddMatch( intSearchID, intScore, intCustID, dateItemDate, strAuthor, strLink, strType )
	With cmdSearch
		.Parameters("@SearchID") = intSearchID
		.Parameters("@Score") = intScore
		.Parameters("@CustomerID") = intCustID
		.Parameters("@Author") = strAuthor
		.Parameters("@Link") = strLink
		.Parameters("@Type") = strType
		.Execute , , adExecuteNoRecords
		intIDOut = .Parameters("@SearchIDOut")
	End With
	AddMatch = intIDOut
End Function

%>