<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then	Redirect("members.asp?Source=admin_buttons_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Configure Your Menu Buttons</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Change Button Setup" then
	intNumMenus = 1

	do until Request("Menu" & intNumMenus) = ""
		intNumMenus = intNumMenus + 1
	loop

	intNumMenus = intNumMenus - 1

	Query = "DELETE MenuButtons WHERE CustomerID = " & CustomerID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "SELECT * FROM MenuButtons"
	Set rsButtons = Server.CreateObject("ADODB.Recordset")
	rsButtons.Open Query, Connect, adOpenStatic, adLockOptimistic

	blCustom = False
	if Request("Custom") = "YES" then blCustom = True

	for i = 1 to intNumMenus
		rsButtons.AddNew

		rsButtons("CustomerID") = CustomerID
		rsButtons("Position") = i
		rsButtons("Name") = Request("Menu" & i)
		rsButtons("Align") = Request("Menu"  & i & "Align")
		rsButtons("Show") = Request("Menu" & i & "Show" )
		rsButtons("CustomLabel") = Request("Menu" & i & "CustomLabel")
		if blCustom then rsButtons("CustomLink") = Request("Menu" & i & "CustomLink" )
		if Request("Menu" & i & "Custom" ) <> "" then rsButtons("Custom") = CInt(Request("Menu" & i & "Custom" ))

		rsButtons.Update

	next


	rsButtons.Update
	rsButtons.Close
	Set rsButtons = Nothing

	Redirect("write_header_footer.asp?Source=admin_buttons_modify.asp?Submit=Changed")

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The button changes have been made.  You can <a href="admin_buttons_modify.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>
<%
'-----------------------Begin Code----------------------------
else
	blParentSiteExists = ParentSiteExists()
	blChildSiteExists = ChildSiteExists()


	Set rsButtons = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM MenuButtons WHERE CustomerID = " & CustomerID & " ORDER BY Position"
	rsButtons.CacheSize = 50
	rsButtons.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	Public intNumButtons, blCustomButtons
	Dim ButtonArray(100, 6)

	blCustomButtons = False

	if rsButtons.EOF then
		intNumButtons = 0

	else
		intNumButtons = rsButtons.RecordCount

		for a = 0 to intNumButtons - 1
			ButtonArray(a, 0) = rsButtons("Name")
			ButtonArray(a, 1) = rsButtons("Align")
			ButtonArray(a, 2) = rsButtons("Show")
			ButtonArray(a, 3) = rsButtons("CustomLabel")
			ButtonArray(a, 4) = rsButtons("CustomLink")
			ButtonArray(a, 5) = rsButtons("Custom")

			if rsButtons("Custom") = 1 then 
				blCustomButtons = True
			end if
			rsButtons.MoveNext
		next
	end if





	Function MustPrintButtonNew( intInclude, strButton, intPos, blOverFlow )
		MustPrintButtonNew = intInclude > 0 and ( ButtonArray(intPos-1, 0) = strButton or (blOverFlow and not InNumButtonArray( strButton )) )
	End Function

	Function InNumButtonArray( strButton )
		for arraylp = 0 to intNumButtons-1
			if ButtonArray(arraylp, 0) = strButton then
				InNumButtonArray = true
				exit function
			end if
		next

		InNumButtonArray = false
	End Function


	Set rsItems = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT ID, Title FROM InfoPages WHERE Title <> 'Home Page' AND CustomerID = " & CustomerID & " AND (ShowButton = 1) ORDER BY Title"
	rsItems.CacheSize = 50
	rsItems.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	Public intMaxPages
	intMaxPages = 0
	if not rsItems.EOF then
		Set PID = rsItems("ID")
		Set PTitle = rsItems("Title")
		Dim strPages(100, 2)
		intMaxPages = rsItems.RecordCount


		for a = 0 to intMaxPages - 1
			strPages(a, 0) = rsItems("ID")
			strPages(a, 1) = rsItems("Title")
			rsItems.MoveNext
		next
	end if
	rsItems.Close




	Public intMaxChildSites
	intMaxChildSites = 0
	if blChildSiteExists then
		Set cmdTemp = Server.CreateObject("ADODB.Command")
		With cmdTemp
			.ActiveConnection = Connect
			.CommandText = "GetChildSitesRecordSet"
			.CommandType = adCmdStoredProc
			.Parameters.Refresh
			.Parameters("@CustomerID") = CustomerID
		End With
		rsItems.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect
		rsItems.CacheSize = 50
		if not rsItems.EOF then
			Dim strChildSites(100, 2)
			intMaxChildSites = rsItems.RecordCount
			for a = 0 to intMaxChildSites - 1
				strChildSites(a, 0) = rsItems("ID")
				strChildSites(a, 1) = rsItems("Title")
				rsItems.MoveNext
			next
		end if
		rsItems.Close

		Set cmdTemp = Nothing
	end if


	intParentSite = 0
	if blParentSiteExists then
		intParentSite = 1
		Set cmdTemp = Server.CreateObject("ADODB.Command")
		With cmdTemp
			.ActiveConnection = Connect
			.CommandText = "GetParentSiteInfo"
			.CommandType = adCmdStoredProc
			.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
			.Parameters.Append .CreateParameter ("@ParentID", adInteger, adParamOutput )
			.Parameters.Append .CreateParameter ("@ShortTitle", adVarWChar, adParamOutput, 100 )
			.Parameters.Append .CreateParameter ("@SubDirectory", adVarWChar, adParamOutput, 100 )
			.Parameters("@CustomerID") = CustomerID
			.Execute , , adExecuteNoRecords
			intParentID = .Parameters("@ParentID")
			strParentShortTitle = .Parameters("@ShortTitle")
		End With
		Set cmdTemp = Nothing
	end if

	set rsItems = Nothing


%>

	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function Changed(form, MasterMenuNum){
			var TestMenuNum, i, blDropout, NumMenus, CurrentSelection, MasterMenu, TestMenu, TestMenuValue, MasterMenuValue, CurrentIndex;

			NumMenus = 1;
			blDropout = false;

			while(!blDropout){
				if (!form.elements['Menu'+NumMenus])
					blDropout = true;
				else
					NumMenus++;

			}

			NumMenus--;

			MasterMenu = form.elements['Menu'+MasterMenuNum];
			CurrentSelection = MasterMenu.selectedIndex;
			MasterMenuValue = MasterMenu.options[CurrentSelection].value;

			for ( TestMenuNum = 1; TestMenuNum <= NumMenus; TestMenuNum++){
				TestMenu = form.elements['Menu'+TestMenuNum];
				TestMenuValue = TestMenu.options[TestMenu.selectedIndex].value;

				if( MasterMenuNum != TestMenuNum ){
					if  ( TestMenuValue == MasterMenuValue ){
						//alert('TestMenuNum-' + TestMenuNum + 'TestMenuValue-' + TestMenuValue + '   MasterMenuNum-' + MasterMenuNum + '   MasterMenuValue-' + MasterMenuValue);

						//This has already been selected.  So now we will change the selection to the first open one
						ChangeSelected( form, TestMenuNum, NumMenus );
						SwitchSelected( form.elements['Menu'+TestMenuNum + 'Show'], form.elements['Menu'+MasterMenuNum + 'Show'] );
						SwitchSelected( form.elements['Menu'+TestMenuNum + 'Align'], form.elements['Menu'+MasterMenuNum + 'Align'] );
						SwitchText( form.elements['Menu'+TestMenuNum + 'CustomLabel'], form.elements['Menu'+MasterMenuNum + 'CustomLabel'] );

						<% if blCustomButtons then %>
						SwitchText( form.elements['Menu'+TestMenuNum + 'CustomLink'], form.elements['Menu'+MasterMenuNum + 'CustomLink'] );
						<% end if %>
					}
				}

			}
	}

	//This function switches two text fields
	function SwitchText( Menu1, Menu2 ){
		var Menu1Selection;

		Menu1Selection = Menu1.value;

		Menu1.value = Menu2.value;
		Menu2.value = Menu1Selection;
	}


	//This function switches two pulldown menus' selections
	function SwitchSelected( Menu1, Menu2 ){
		var Menu1Selection;

		Menu1Selection = Menu1.selectedIndex;

		Menu1.selectedIndex = Menu2.selectedIndex;
		Menu2.selectedIndex = Menu1Selection;
	}


	//This function changes the value of a pulldown to the first open value
	function ChangeSelected( form, MasterMenuNum, NumMenus ){
		var form, blDropout, NumMenus, CurrentMasterValue, blOptionTaken, blDone, TestMenuNum, i, CurrentSelection, MasterMenu, TestMenu, TestMenuValue, MasterMenuValue, CurrentIndex, MasterMenuLength;


		MasterMenu = form.elements['Menu'+MasterMenuNum];
		CurrentSelection = MasterMenu.selectedIndex;
		MasterMenuValue = MasterMenu.options[CurrentSelection].value;

		blDone = false;

		//go through the options and find the newest open one

		MenuIndex = MasterMenu.length;

		//OUTTER LOOP.  Loop through options in the master menu, skipping the selected one
		//INNER LOOP:  If the option is not taken by any menu other than the master, then that is the only available option
		//             which is the one we will change to.  This should only happen ONCE


		for ( i = 0; i < MenuIndex; i++ ){
			if ( i != CurrentSelection ){
				blOptionTaken = false; 

				CurrentMasterValue = MasterMenu.options[i].value;


				for ( TestMenuNum = 1; TestMenuNum <= NumMenus; TestMenuNum++){
					if ( TestMenuNum != MasterMenuNum ){
						TestMenu = form.elements['Menu'+TestMenuNum];
						TestMenuValue = TestMenu.options[TestMenu.selectedIndex].value;

						if ( TestMenuValue == CurrentMasterValue )
							blOptionTaken = true;

						//alert('blOptionTaken = ' + blOptionTaken + ' OptionIndex = ' + i + ' MasterMenuNum = ' + MasterMenuNum + '  TestMenuNum = ' + TestMenuNum + 'TestMenuValue = ' + TestMenuValue + '  MasterValue = ' + MasterMenuValue + '  CuurentMasterValue = ' + CurrentMasterValue );
					}
				}

				//This option is open.  select it and exit the loop
				if ( !blOptionTaken )
					MasterMenu.selectedIndex = i;

			}
		}
	}


	//-->
	</SCRIPT>


	<p>You can easily change the order and placement of your buttons below.</p>
	<form METHOD="post" ACTION="admin_buttons_modify.asp" name="myForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
	<tr>
		<td class="TDHeader" colspan=2 align="center">Button Order</td>
		<td class="TDHeader">Page Position</td>
		<td class="TDHeader">Shown In</td>
		<td class="TDHeader">Custom Label</td>
		<% if blCustomButtons then %>
			<input type="hidden" name="Custom" value="YES">
		<td class="TDHeader">Custom Link</td>
		<% end if %>
	</tr>

<%
	blOverflow = false

	Public intMenuNum
	intMenuNum = 0


	for i = 1 to intNumButtons + 1

		if i = intNumButtons + 1 then blOverflow = true

		if MustPrintButtonNew( 1, "Home", i, blOverFlow ) then PrintPullDown i, "Home"

		if blParentSiteExists and MustPrintButtonNew( 1, "Parent"&intParentID, i, blOverFlow ) then PrintPullDown i, "Parent"&intParentID

		if blChildSiteExists then
			for p = 0 to intMaxChildSites - 1
				if MustPrintButtonNew( 1, "Child"&strChildSites(p, 0), i, blOverFlow ) then PrintPullDown i, "Child"&strChildSites(p, 0)
			next
		end if


		for p = 0 to intMaxPages - 1
			if MustPrintButtonNew( 1, "InfoPage"&strPages(p, 0), i, blOverFlow ) then PrintPullDown i, "InfoPage"&strPages(p, 0)
		next

		if ButtonArray(i-1, 5) = 1 then
			if MustPrintButtonNew( 1, ButtonArray(i-1, 0), i, blOverFlow ) then PrintPullDown i, ButtonArray(i-1, 0)
		end if

		if MustPrintButtonNew( IncludeAnnouncements, "Announcements", i, blOverFlow ) then PrintPullDown i, "Announcements"
		if MustPrintButtonNew( IncludeMeetings, "Meetings", i, blOverFlow ) then PrintPullDown i, "Meetings"
		if MustPrintButtonNew( IncludeStories, "Stories", i, blOverFlow ) then PrintPullDown i, "Stories"
		if MustPrintButtonNew( IncludeCalendar, "Calendar", i, blOverFlow ) then PrintPullDown i, "Calendar"
		if MustPrintButtonNew( IncludeLinks, "Links", i, blOverFlow ) then PrintPullDown i, "Links"
		if MustPrintButtonNew( IncludeQuotes, "Quotes", i, blOverFlow ) then PrintPullDown i, "Quotes"
		if MustPrintButtonNew( IncludeGuestbook, "Guestbook", i, blOverFlow ) then PrintPullDown i, "Guestbook"
		if MustPrintButtonNew( IncludeForum, "Forum", i, blOverFlow ) then PrintPullDown i, "Forum"
		if MustPrintButtonNew( IncludePhotos, "Photos", i, blOverFlow ) then PrintPullDown i, "Photos"
		if MustPrintButtonNew( IncludeVoting, "Voting", i, blOverFlow ) then PrintPullDown i, "Voting"
		if MustPrintButtonNew( IncludeQuizzes, "Quizzes", i, blOverFlow ) then PrintPullDown i, "Quizzes"
		if MustPrintButtonNew( IncludeMedia, "Media", i, blOverFlow ) then PrintPullDown i, "Media"
		if MustPrintButtonNew( IncludeNewsletter, "Newsletter", i, blOverFlow ) then PrintPullDown i, "Newsletter"
		if MustPrintButtonNew( IncludeStore + AllowStore - 1, "Store", i, blOverFlow ) then PrintPullDown i, "Store"

		if MustPrintButtonNew( IncludeStats, "Stats", i, blOverFlow ) then PrintPullDown i, "Stats"
		if MustPrintButtonNew( 1, "Members", i, blOverFlow ) then PrintPullDown i, "Members"


		if MustPrintButtonNew( 1, "Search", i, blOverFlow ) then PrintPullDown i, "Search"

		if not rsButtons.EOF then rsButtons.MoveNext

	next


	Set rsButtons = Nothing



%>

		<tr>
    		<td colspan="5" align="center" class="TDMain1">


				<input type="submit" name="Submit" value="Change Button Setup">
	   		</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if



	Sub PrintPullDown ( intButtonNum, strSelect )

		intMenuNum = intMenuNum + 1
		if intMenuNum < 10 then
			strNum = "&nbsp;&nbsp;" & intMenuNum
		else
			strNum = intMenuNum
		end if
%>
		<tr><td class="<% PrintTDMain %>" align="right"><%=intMenuNum%>.</td>

		<td class="<% PrintTDMain %>"><select name="Menu<%=intMenuNum%>" onChange="Changed(this.form, <%=intMenuNum%>)">
<%
		blTestOver = false

		for j = 1 to intNumButtons + 1
			if j = intNumButtons + 1 then blTestOver = true

			if MustPrintButtonNew( 1, "Home", j, blTestOver ) then WriteOption "Home", "Home", strSelect

			if blParentSiteExists and MustPrintButtonNew( 1, "Parent"&intParentID, j, blTestOver ) then WriteOption "Parent"&intParentID, strParentShortTitle, strSelect


			if blChildSiteExists then
				for k = 0 to intMaxChildSites - 1
					if MustPrintButtonNew( 1, "Child"&strChildSites(k, 0), j, blTestOver ) then WriteOption "Child"&strChildSites(k, 0), strChildSites(k, 1), strSelect
				next
			end if


			for k = 0 to intMaxPages - 1
				if MustPrintButtonNew( 1, "InfoPage"&strPages(k, 0), j, blTestOver ) then WriteOption "InfoPage"&strPages(k, 0),  strPages(k, 1), strSelect
			next

		if ButtonArray(j-1, 5) = 1 then
			if MustPrintButtonNew( 1, ButtonArray(j-1, 0), j, blOverFlow ) then WriteOption ButtonArray(j-1, 3), ButtonArray(j-1, 0), strSelect
		end if

			if MustPrintButtonNew( IncludeAnnouncements, "Announcements", j, blTestOver ) then WriteOption "Announcements", AnnouncementsTitle, strSelect
			if MustPrintButtonNew( IncludeMeetings, "Meetings", j, blTestOver ) then WriteOption "Meetings", MeetingsTitle, strSelect
			if MustPrintButtonNew( IncludeStories, "Stories", j, blTestOver ) then WriteOption "Stories", StoriesTitle, strSelect
			if MustPrintButtonNew( IncludeCalendar, "Calendar", j, blTestOver ) then WriteOption "Calendar", CalendarTitle, strSelect
			if MustPrintButtonNew( IncludeLinks, "Links", j, blTestOver ) then WriteOption "Links", LinksTitle, strSelect
			if MustPrintButtonNew( IncludeQuotes, "Quotes", j, blTestOver ) then WriteOption "Quotes", QuotesTitle, strSelect
			if MustPrintButtonNew( IncludeGuestbook, "Guestbook", j, blTestOver ) then WriteOption "Guestbook", GuestbookTitle, strSelect
			if MustPrintButtonNew( IncludeForum, "Forum", j, blTestOver ) then WriteOption "Forum", ForumTitle, strSelect
			if MustPrintButtonNew( IncludePhotos, "Photos", j, blTestOver ) then WriteOption "Photos", PhotosTitle, strSelect
			if MustPrintButtonNew( IncludeVoting, "Voting", j, blTestOver ) then WriteOption "Voting", VotingTitle, strSelect
			if MustPrintButtonNew( IncludeQuizzes, "Quizzes", j, blTestOver ) then WriteOption "Quizzes", QuizzesTitle, strSelect
			if MustPrintButtonNew( IncludeMedia, "Media", j, blTestOver ) then WriteOption "Media", MediaTitle, strSelect
			if MustPrintButtonNew( IncludeNewsletter, "Newsletter", j, blTestOver ) then WriteOption "Newsletter", NewsletterTitle, strSelect
			if MustPrintButtonNew( IncludeStore + AllowStore - 1, "Store", j, blTestOver ) then WriteOption "Store", StoreTitle, strSelect

			'Custom sub for the buttons
			if MustPrintButtonNew( 1, "Custom", j, blTestOver ) then WriteOption "Custom", "Custom Buttons", strSelect
			if MustPrintButtonNew( IncludeStats, "Stats", j, blTestOver ) then WriteOption "Stats", StatsTitle, strSelect
			if MustPrintButtonNew( 1, "Members", j, blTestOver ) then WriteOption "Members", MembersTitle, strSelect

			if MustPrintButtonNew( 1, "Search", j, blTestOver ) then WriteOption "Search", "Search", strSelect
		next
%>
		</select>
		</td><td class="<% PrintTDMain %>">
	
		<input type="hidden" name="Menu<%=intMenuNum%>Custom" value="<%=ButtonArray(intButtonNum-1, 5)%>">
		<select name="Menu<%=intMenuNum%>Align" onChange="Changed(this.form, <%=intMenuNum%>)">
<%

		WriteOption "Left", "Left", ButtonArray(intButtonNum-1, 1)
		WriteOption "Right", "Right", ButtonArray(intButtonNum-1, 1)
		WriteOption "Top", "Top", ButtonArray(intButtonNum-1, 1)
%>
		</select>
		</td><td class="<% PrintTDMain %>">
		<select name="Menu<%=intMenuNum%>Show" onChange="Changed(this.form, <%=intMenuNum%>)">
<%
		WriteOption "Menu", "Main Menu Only", ButtonArray(intButtonNum-1, 2)
		WriteOption "Footer", "Footer Only", ButtonArray(intButtonNum-1, 2)
		WriteOption "MenuFooter", "Main Menu And Footer", ButtonArray(intButtonNum-1, 2)
		WriteOption "Nowhere", "Do Not Show Button", ButtonArray(intButtonNum-1, 2)
%>
		</select></td>
		<td class="<% PrintTDMain %>">
		<input type="text" size="15" name="Menu<%=intMenuNum%>CustomLabel" onChange="Changed(this.form, <%=intMenuNum%>)" value="<%=ButtonArray(intButtonNum-1, 3)%>">
		</td>		
		<% if blCustomButtons then %>		
		<td class="<% PrintTDMain %>">
		<input type="text" size="15" name="Menu<%=intMenuNum%>CustomLink" onChange="Changed(this.form, <%=intMenuNum%>)" value="<%=ButtonArray(intButtonNum-1, 4)%>">
		</td>	
		<% end if %>
		</tr>
<%
	End Sub
%>
