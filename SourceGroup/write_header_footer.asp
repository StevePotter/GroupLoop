
<%
'This script takes all the information from a site and generates the header and footer files for it
'I used to have a header and footer.asp files, but I figure this is much faster, especially when
'lots of people are requesting pages at once.  This saves like 80 if thens.  This script is 
'executed when scripts that could change the header are executed.  It has a Request("Source")
'variable sent, so they can be redirected once the header and footer files are changed.
'First we are going to do the header, then the footer files

'-----------------------Create Files----------------------------
'if the constants aren't defined in the const.inc file, get the ones we need so it doesn't fuck up
'this is very necessary for when the site first gets created
if CellSpacing = "" or IncludeAnnouncements = "" then
	'Open up the configuration recordset
	if CustomerID = "" then CustomerID = intCustomerID
	Query = "SELECT * FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

	'Get the subdirectory from the customer record
	Set Command = Server.CreateObject("ADODB.Command")
	With Command
		'Check the scheme to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandType = adCmdStoredProc
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		strSubDir = .Parameters("@SubDirectory")
	End With
	Set Command = Nothing

	CellSpacing = rsConfig("CellSpacing")
	CellPadding = rsConfig("CellPadding")
	Border = rsConfig("Border")
	HeadingAlignment = rsConfig("HeadingAlignment")
	PageSize = rsConfig("PageSize")
	IncludeStories = rsConfig("IncludeStories")
	IncludeAnnouncements = rsConfig("IncludeAnnouncements")
	IncludeCalendar = rsConfig("IncludeCalendar")
	IncludeQuizzes = rsConfig("IncludeQuizzes")
	IncludeVoting = rsConfig("IncludeVoting")
	IncludePhotos = rsConfig("IncludePhotos")
	IncludePhotoCaptions = rsConfig("IncludePhotoCaptions")
	IncludeGuestbook = rsConfig("IncludeGuestbook")
	IncludeForum = rsConfig("IncludeForum")
	IncludeStats = rsConfig("IncludeStats")
	IncludeLinks = rsConfig("IncludeLinks")
	IncludeAdditions = rsConfig("IncludeAdditions")
	IncludeQuotes = rsConfig("IncludeQuotes")
	IncludeMedia = rsConfig("IncludeMedia")
	IncludeNewsletter = rsConfig("IncludeNewsletter")
	IncludeMeetings = rsConfig("IncludeMeetings")
	IncludeStore = rsConfig("IncludeStore")
	AllowStore = rsConfig("AllowStore")
	Title = rsConfig("Title")
	AdditionsTitle = rsConfig("AdditionsTitle")
	AnnouncementsTitle = rsConfig("AnnouncementsTitle")
	StoriesTitle = rsConfig("StoriesTitle")
	CalendarTitle = rsConfig("CalendarTitle")
	QuizzesTitle = rsConfig("QuizzesTitle")
	VotingTitle = rsConfig("VotingTitle")
	PhotosTitle = rsConfig("PhotosTitle")
	PhotoCaptionsTitle = rsConfig("PhotoCaptionsTitle")
	ForumTitle = rsConfig("ForumTitle")
	MembersTitle = rsConfig("MembersTitle")
	GuestbookTitle = rsConfig("GuestbookTitle")
	LinksTitle = rsConfig("LinksTitle")
	NewsTitle = rsConfig("NewsTitle")
	AboutSiteTitle = rsConfig("AboutSiteTitle")
	StatsTitle = rsConfig("StatsTitle")
	QuotesTitle = rsConfig("QuotesTitle")
	MediaTitle = rsConfig("MediaTitle")
	NewsletterTitle = rsConfig("NewsletterTitle")
	StoreTitle = rsConfig("StoreTitle")
	MeetingsTitle = rsConfig("MeetingsTitle")
	PhotosMegs = rsConfig("PhotosMegs")
	IncludeEditSectionPropButtons = rsConfig("IncludeEditSectionPropButtons")
	Subdirectory = strSubDir
	'Put the https address if they choose to use secure logins
	if rsConfig("SecureLogin") = 0 then
		SecurePath = "http://www.GroupLoop.com/" & strSubDir & "/"
	else
		SecurePath = "https://www.OurClubPage.com/" & strSubDir & "/"
	end if
	NonSecurePath = "http://www.GroupLoop.com/" & strSubDir & "/"

	rsConfig.Close
	set rsConfig = Nothing

'	Dim Buttons(100, 3)
end if

'if not IsObject(Buttons) then 

'Open up the images recordset
Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
Set rsLook = Server.CreateObject("ADODB.Recordset")
rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly

'Don't touch their shit
if rsLook("CustomHeader") = 1 then
	rsLook.Close
	Set rsLook = Nothing
	if Request("Source") <> "" then
		Redirect Request("Source")
	else
		Redirect "index.asp"
	end if
end if

if strPath = "" then strPath = GetPath("")
Set FileSystem = CreateObject("Scripting.FileSystemObject")
if FileSystem.FileExists (strPath & "header.asp") then FileSystem.DeleteFile (strPath & "header.asp")
if FileSystem.FileExists (strPath & "footer.asp") then FileSystem.DeleteFile (strPath & "footer.asp")
Set HeaderFile = FileSystem.CreateTextFile(strPath & "header.asp")
Set FooterFile = FileSystem.CreateTextFile(strPath & "footer.asp")

strImagePath = GetPath("images")

Set rsPages = Server.CreateObject("ADODB.Recordset")

'-----------------------Start Writing Files----------------------------

'-----------------------Start Writing Head----------------------------

HeaderFile.Write "<html>"
HeaderFile.Write "<head>"
HeaderFile.Write "<title>" & Title & "</title>"
HeaderFile.Write "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">"
HeaderFile.Write "<meta name=" & Chr(34) & "Author" & Chr(34) & " content=" & Chr(34) & "GroupLoop.com - Get Your Own Interactive Community!" & Chr(34) & ">"
HeaderFile.Write "<meta name=" & Chr(34) & "KeyWords" & Chr(34) & " content=" & Chr(34) & rsLook("Keywords") & Chr(34) & ">"
HeaderFile.Write "<meta name=" & Chr(34) & "Description" & Chr(34) & " content=" & Chr(34) & rsLook("Description") & Chr(34) & ">"


Function GetItalic( intItalic )
	if intItalic = 1 then
		GetItalic = "; font-style: italic"
	else
		GetItalic = ""
	end if

End Function
Function GetBold( intBold )
	if intBold = 1 then
		GetBold = "; font-weight: bold"
	else
		GetBold = ""
	end if

End Function



Function GetLinkUnderline( intUnderline )
	if intUnderline = 1 then
		GetLinkUnderline = ""
	else
		GetLinkUnderline = "; text-decoration:none"
	end if

End Function

'-----------------------Start Writing Styles----------------------------
	HeaderFile.WriteBlankLines 2
	HeaderFile.Write "<style type=" & Chr(34) & "text/css" & Chr(34) & ">"

	strBold = GetBold(rsLook("BodyTextBold"))
	strItalic = GetItalic(rsLook("BodyTextItalic"))

	strBack = ""
	if rsLook("BackgroundColor") <> "" then strBack = "; background-color: " & rsLook("BackgroundColor") & " "

	HeaderFile.Write "BODY{font-family:" & rsLook("BodyTextFont") & "; font-size:" & rsLook("BodyTextSize") & "px; color: " & rsLook("BodyTextColor") & strBack & strBold & strItalic & "}"
	HeaderFile.Write ".BodyText{font-family:" & rsLook("BodyTextFont") & "; font-size:" & rsLook("BodyTextSize") & "px; color: " & rsLook("BodyTextColor") & strBack & strBold & strItalic & "}"
	HeaderFile.Write ".LinkText{font-family:" & rsLook("BodyTextFont") & "; font-size:" & rsLook("LinkSize") & "px; color: " & rsLook("LinkColor") & strBold & strItalic & GetLinkUnderline(rsLook("LinkUnderline")) & "}"
	HeaderFile.Write "form {margin-bottom : 0; }"


	HeaderFile.Write ".LeftMenu{font-family:" & rsLook("LeftMenuFont") & "; font-size:" & rsLook("LeftMenuSize") & "px; color: " & rsLook("LeftMenuColor") & GetBold(rsLook("LeftMenuBold")) & GetItalic(rsLook("LeftMenuItalic")) & GetLinkUnderline(rsLook("LeftMenuUnderline")) & "}"
	HeaderFile.Write ".RightMenu{font-family:" & rsLook("RightMenuFont") & "; font-size:" & rsLook("RightMenuSize") & "px; color: " & rsLook("RightMenuColor") & GetBold(rsLook("RightMenuBold")) & GetItalic(rsLook("RightMenuItalic")) & GetLinkUnderline(rsLook("RightMenuUnderline")) & "}"
	HeaderFile.Write ".TopMenu{font-family:" & rsLook("TopMenuFont") & "; font-size:" & rsLook("TopMenuSize") & "px; color: " & rsLook("TopMenuColor") & GetBold(rsLook("TopMenuBold")) & GetItalic(rsLook("TopMenuItalic")) & GetLinkUnderline(rsLook("TopMenuUnderline")) & "}"

	HeaderFile.Write ".Heading{font-family:" & rsLook("HeadingFont") & "; font-size:" & rsLook("HeadingSize") & "px; color: " & rsLook("HeadingColor") & GetBold(rsLook("HeadingBold")) & GetItalic(rsLook("HeadingItalic")) & "}"
	HeaderFile.Write ".Title{font-family:" & rsLook("TitleFont") & "; font-size:" & rsLook("TitleSize") & "px; color: " & rsLook("TitleColor") & GetBold(rsLook("TitleBold")) & GetItalic(rsLook("TitleItalic")) & "}"

	strBold = GetBold(rsLook("TableMainTextBold"))
	strItalic = GetItalic(rsLook("TableMainTextItalic"))

	strBack1 = ""
	strBack2 = ""
	if rsLook("TableMainBackground1") <> "" then strBack1 = "; background-color: " & rsLook("TableMainBackground1") & " "
	if rsLook("TableMainBackground2") <> "" then strBack2 = "; background-color: " & rsLook("TableMainBackground2") & " "

	HeaderFile.Write ".TDMain1{font-family:" & rsLook("TableMainTextFont") & "; font-size:" & rsLook("TableMainTextSize") & "px; color: " & rsLook("TableMainTextColor") & strBack1 & strBold & strItalic & "}"
	HeaderFile.Write ".TDMain2{font-family:" & rsLook("TableMainTextFont") & "; font-size:" & rsLook("TableMainTextSize") & "px; color: " & rsLook("TableMainTextColor") & strBack2 & strBold & strItalic & "}"

	if rsLook("TableHeaderBackground") <> "" then strBack = "; background-color: " & rsLook("TableHeaderBackground") & " "
	HeaderFile.Write ".TDHeader{font-family:" & rsLook("TableHeaderTextFont") & "; font-size:" & rsLook("TableHeaderTextSize") & "px; color: " & rsLook("TableHeaderTextColor") & strBack & GetBold(rsLook("TableHeaderTextBold")) & GetItalic(rsLook("TableHeaderTextItalic")) & "}"

	HeaderFile.Write "</style>"
	HeaderFile.WriteBlankLines 2

'-----------------------End Writing Styles----------------------------



HeaderFile.Write "<script language=" &Chr(34)& "JavaScript1.2" &Chr(34)& " src=" &Chr(34)& "scripts.js" &Chr(34)& " type=" &Chr(34)& "text/javascript" &Chr(34)& "></script>" & _
				"<!-- #include file=" & Chr(34) & "constants.inc" & Chr(34) & " -->" & _

				"<" & "% if (not LoggedMember and SiteMembersOnly) and not blBypass then Redirect " & Chr(34) & "login.asp?Type=Master" & Chr(34) & " %" & ">" & _
				"</head>"

HeaderFile.WriteBlankLines 2

'Put some site-specific variables into the browser
HeaderFile.WriteLine "<SCRIPT LANGUAGE=" & Chr(34) & "JavaScript" & Chr(34) & ">"
HeaderFile.WriteLine "<!--"
HeaderFile.WriteLine "var ShowSectionPropButtons;"
HeaderFile.WriteLine "ShowSectionPropButtons = " & IncludeEditSectionPropButtons & ";"
HeaderFile.WriteLine "//-->"
HeaderFile.WriteLine "</SCRIPT>"


'If they use a background image, then put it in.  Else use the specified color
if ImageExists("BackgroundImage", strExt) then
	strBackground = "background='images/BackgroundImage." & strExt & "'"
else
	strBackground = "bgcolor='"& rsLook("BackgroundColor") & "'"
end if


'All the preload shit
strPreload = GetCustomPreload
if ImageExists( "TitleRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "TitleRolloverImage" & "." & strExt & "',"
if ImageExists( "HomeRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "HomeRolloverImage" & "." & strExt & "',"
if ImageExists( "CalendarRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "CalendarRolloverImage" & "." & strExt & "',"
if ImageExists( "AboutSiteRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "AboutSiteRolloverImage" & "." & strExt & "',"
if ImageExists( "AnnouncementsRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "AnnouncementsRolloverImage" & "." & strExt & "',"
if ImageExists( "MeetingsRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "MeetingsRolloverImage" & "." & strExt & "',"
if ImageExists( "StoriesRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "StoriesRolloverImage" & "." & strExt & "',"
if ImageExists( "ForumRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "ForumRolloverImage" & "." & strExt & "',"
if ImageExists( "PhotosRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "PhotosRolloverImage" & "." & strExt & "',"
if ImageExists( "QuizzesRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "QuizzesRolloverImage" & "." & strExt & "',"
if ImageExists( "VotingRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "VotingRolloverImage" & "." & strExt & "',"
if ImageExists( "LinksRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "LinksRolloverImage" & "." & strExt & "',"
if ImageExists( "GuestbookRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "GuestbookRolloverImage" & "." & strExt & "',"
if ImageExists( "StatsRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "StatsRolloverImage" & "." & strExt & "',"
if ImageExists( "QuotesRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "QuotesRolloverImage" & "." & strExt & "',"
if ImageExists( "MediaRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "MediaRolloverImage" & "." & strExt & "',"
if ImageExists( "MeetingRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "MeetingRolloverImage" & "." & strExt & "',"
if ImageExists( "NewsletterRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "NewsletterRolloverImage" & "." & strExt & "',"
if ImageExists( "StoreRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "StoreRolloverImage" & "." & strExt & "',"
if ImageExists( "ChatRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "ChatRolloverImage" & "." & strExt & "',"
if ImageExists( "MembersRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "MembersRolloverImage" & "." & strExt & "',"
if ImageExists( "SearchRolloverImage", strExt ) then strPreload = strPreload & "'images/" & "SearchRolloverImage" & "." & strExt & "',"

if strPreload <> "" then strPreload = Left( strPreload, (Len(strPreload) - 1) )

'Write the body tag
HeaderFile.Write "<body " & strBackground & " text=" & rsLook("BodyTextColor") & " link=" & rsLook("LinkColor") & " vlink=" & rsLook("VisitedLinkColor") & " "
'Write the preload shit
if strPreload <> "" then HeaderFile.Write " onLoad=" & Chr(34) & "MM_preloadImages(" & strPreload & ");" & Chr(34)
'End the body tag
HeaderFile.Write ">"


blAlreadyOutput = False


'The main page table heading.  this lets us easily control the total width of the page
HeaderFile.Write "<table width=" & Chr(34) & rsLook("TotalPageWidth") & Chr(34) & " align=" & rsLook("TotalPageAlignment") & " border='0' cellspacing='0' cellpadding='0'><tr><td>"

'There are two page layouts

'        LAYOUT 1                       LAYOUT 2
' _____________________          _____________________
' |                   |          |    |         |    |
' |________Top________|          |    |___Top___|    |
' |    |         |    |          |    |         |    |
' |    |         |    |          |    |         |    |
' | Lft|         | Rt |          | Lft|         | Rt |
' |    |         |    |          |    |         |    |
' |    |         |    |          |    |         |    |
' |    |         |    |          |    |         |    |
' |    |         |    |          |    |         |    |
' |    |         |    |          |    |         |    |
' |____|_________|____|          |____|_________|____|
'

Function GetMenuBackground(strMenuAlign)
	strMenuBackground = ""
	strFileName = strMenuAlign & "MenuBackgroundImage"

	if ImageExists( strFileName, strExt ) then strMenuBackground = " background='images/" & strFileName & "." & strExt & "' "

	GetMenuBackground = strMenuBackground


End Function


Public intNumButtons, intMaxPages, intMaxChildSites, blParentSiteExists, blChildSiteExists, blButtons

blParentSiteExists = ParentSiteExists()
blChildSiteExists = ChildSiteExists()

Dim ButtonArray(100, 6)
Dim InfoPages(100, 2)
Dim ChildSites(100, 3)
Dim ParentSites(1, 3)


FillArrays

'Layout 1 - the top menu spans the whole page
if not cBool(rsLook("TopMenuShare")) then
	PrintTop
	PrintLeft
else
	PrintLeft
	PrintTop
end if

PrintRight

PrintBottom

'This prints the title and possible buttons across the top of the screen
Sub PrintTop()
	blTopMenuAboveTitle = CBool(rsLook("TopMenuAboveTitle"))

	if not blTopMenuAboveTitle then PrintTitle

	blFullLength = CBool(rsLook("TopMenuFullBackground"))

	'If we have buttons up top print them out
	if SectorHasButtons("Top") then

		'Print out the title
		HeaderFile.Write "<table width=" & Chr(34) & "100%" & Chr(34) & " border=0 cellspacing=0 cellpadding=0>"

		'Start printing out the menu
		HeaderFile.Write "<tr>"

		PrintImage blImageExists, "TopMenuTop", "TopMenuTop", HeaderFile, "<td width=1>", "</td>"

		HeaderFile.Write "<td class=LinkText valign=" & rsLook("TopMenuVAlignment") & _
		" align=" & rsLook("TopMenuAlignment") & "  " & GetMenuBackground("Top") & ">"

		strSeparator = GetSeparator("Top")

		PrintButtons HeaderFile, strSeparator, "Top", False


		'End this cell
		HeaderFile.Write "</td>"

		PrintImage blImageExists, "TopMenuBottom", "TopMenuBottom", HeaderFile, rsLook("TopMenuSeparator") & "<td width=1>", "</td>"

		HeaderFile.Write "</tr></table>"

	end if


	if blTopMenuAboveTitle then PrintTitle

End Sub



Sub PrintTitle()


	'Get the page title
	if ImageExists( "TitleImage", strExt ) and not ImageExists( "TitleRolloverImage", strOverExt ) then
		strTitle = "<img src='images/TitleImage." & strExt & "' border='0' alt='" & Title & "'>"
	elseif ImageExists( "TitleImage", strExt ) and ImageExists( "TitleRolloverImage", strOverExt ) then
		strTitle = "<a href=" & Chr(34) & NonSecurePath & "index.asp?Action=Old" & Chr(34) & " onMouseOver=" & Chr(34) & "document.images['Title'].src='images/TitleRollover"&"Image." & strOverExt & "'" & Chr(34) & " onMouseOut=" & Chr(34) & "document.images['Title'].src='images/Title"&"Image." & strExt & "'" & Chr(34) & "><img src='images/Title"&"Image." & strExt & "' alt='" & Title & "' border=0 name='Title'></a>"
	else
		strTitle = Title
	end if

	'Should we put a space
	if rsLook("TitleSpace") = 1 then strTitle = strTitle & "<br><br>"

	'Print out the title
	HeaderFile.Write "<table width=" & Chr(34) & "100%" & Chr(34) & " border=0 cellspacing=0 cellpadding=0><tr>"

	PrintImage blImageExists, "TitleTop", "TitleTop", HeaderFile, "<td width=1>", "</td>"

	HeaderFile.Write "<td class=Title valign=" & rsLook("TitleVAlignment") & _
	" align=" & rsLook("TitleAlignment") & "  " & GetMenuBackground("Title") & ">" & _
		strTitle & "</td>"

	PrintImage blImageExists, "TitleBottom", "TitleBottom", HeaderFile, "<td width=1>", "</td>"

	HeaderFile.Write "</tr></table>"



End Sub


'This prints the menu across the left of the screen
Sub PrintLeft()

	'Open up the table that will encapsulate the left menu, body, and right menu, and possibly the top menu
	'Print out the title
	HeaderFile.Write "<table width=" & Chr(34) & "100%" & Chr(34) & " border=0 cellspacing=0 cellpadding=0><tr>"

	blFullLength = CBool(rsLook("LeftMenuFullBackground"))

	'If we have buttons to print
	if SectorHasButtons("Left") or not blButtons then
		'Open up the column in the table for the buttons

		if blFullLength then
			HeaderFile.Write "<td class=LinkText valign=" & rsLook("LeftMenuVAlignment") & GetMenuWidth("Left") & _
			" align=" & rsLook("LeftMenuAlignment") &  "  " & GetMenuBackground("Left") & ">"

		else
			HeaderFile.Write "<td class=LinkText valign=" & rsLook("LeftMenuVAlignment") & GetMenuWidth("Left") & _
			" align=" & rsLook("LeftMenuAlignment") & ">"

			HeaderFile.Write "<table width=" & Chr(34) & "100%" & Chr(34) & " border=0 cellspacing=0 cellpadding=0><tr>"

			HeaderFile.Write "<td class=LinkText valign=" & rsLook("LeftMenuVAlignment") & GetMenuWidth("Left") & _
			" align=" & rsLook("LeftMenuAlignment") & "  " & GetMenuBackground("Left") & ">"
		end if

		strSeparator = GetSeparator("Left")


		PrintButtons HeaderFile, strSeparator, "Left", True

		'End this cell
		if not blFullLength then HeaderFile.Write "</td></tr></table>"
		HeaderFile.Write "</td>"
	end if

	'Now print out the heading for the cell containing the body
	'Open up the column in the table for the buttons
	HeaderFile.Write "<td valign=" & rsLook("BodyVAlignment") & _
	" align=" & rsLook("BodyAlignment") & "  " & GetMenuBackground("Body") & ">"

End Sub


'This prints the menu across the left of the screen
Sub PrintRight()

	blFullLength = CBool(rsLook("RightMenuFullBackground"))


	'If we have buttons to print
	if SectorHasButtons("Right") then
		'Open up the column in the table for the buttons


		if blFullLength then
			FooterFile.Write "<td class=LinkText valign=" & rsLook("RightMenuVAlignment") & GetMenuWidth("Right") & _
			" align=" & rsLook("RightMenuAlignment") &  "  " & GetMenuBackground("Right") & ">"

		else
			FooterFile.Write "<td class=LinkText valign=" & rsLook("RightMenuVAlignment") & GetMenuWidth("Right") & _
			" align=" & rsLook("RightMenuAlignment") & ">"

			FooterFile.Write "<table width=" & Chr(34) & "100%" & Chr(34) & " border=0 cellspacing=0 cellpadding=0><tr>"

			FooterFile.Write "<td class=LinkText valign=" & rsLook("RightMenuVAlignment") & GetMenuWidth("Right") & _
			" align=" & rsLook("RightMenuAlignment") & "  " & GetMenuBackground("Right") & ">"
		end if

		strSeparator = GetSeparator("Right")

		PrintButtons FooterFile, strSeparator, "Right", True

		'End this cell
		if not blFullLength then FooterFile.Write "</td></tr></table>"
		FooterFile.Write "</td>"

		'End this cell
		FooterFile.Write "</td>"
	end if

	'Now close out the main page table
	FooterFile.Write "</td></table>"

End Sub

Sub PrintBottom()
	FooterFile.Write "<br>"

	'Show the footer
	if rsLook("ShowFooter") = 1 then
		FooterFile.Write "<br><table width=" & Chr(34) & "100%" & Chr(34) & " cellspacing=2 cellpadding=1 border=0><tr><td class=LinkText align=" & rsLook("FooterAlignment") & ">"

		if rsLook("FooterSource") <> "" then FooterFile.Write "<span class=BodyText>" & rsLook("FooterSource") & "</span><br>"



		blOverflow = false


		blFirst = true

		for i = 1 to intNumButtons + 1
			if i = intNumButtons + 1 then blOverflow = true

			if MustPrintFooterButton( 1, "Home", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & "index.asp?Action=Old" & Chr(34) & ">Home</a>]  "



		if blParentSiteExists then
			if MustPrintFooterButton( 1, "Parent"&ParentSites(0, 0), i, blOverFlow ) then
					FooterFile.Write "[<a href='http://www.GroupLoop.com/" & ParentSites(0, 2) & "'>" & GetButtonTitle( ParentSites(0, 1), "Parent"&ParentSites(0, 0) ) & "</a>]  "
			end if
		end if

		if blChildSiteExists then
			for p = 0 to intMaxChildSites - 1
				if MustPrintFooterButton( 1, "Child"&ChildSites(p, 0), i, blOverFlow ) then
					FooterFile.Write "[<a href='http://www.GroupLoop.com/" & ChildSites(p, 2) & "'>" & GetButtonTitle( ChildSites(p, 1), "Child"&ChildSites(p, 0) ) & "</a>]  "
				end if
			next
		end if

		for p = 0 to intMaxPages - 1
			if MustPrintFooterButton( 1, "InfoPage"&InfoPages(p, 0), i, blOverFlow ) then
					FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & "pages_read.asp?ID=" & InfoPages(p, 0) & Chr(34) & ">" & GetButtonTitle( InfoPages(p, 1), "InfoPage"&InfoPages(p, 0) ) & "</a>]  "
			end if
		next

			if MustPrintFooterButton( IncludeAnnouncements, "Announcements", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "announcements.asp", "Announcements" ) & Chr(34) & ">" & GetButtonTitle( AnnouncementsTitle, "Announcements" ) & "</a>]  "
			if MustPrintFooterButton( IncludeMeetings, "Meetings", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "meetings.asp", "Meetings" ) & Chr(34) & ">" &  GetButtonTitle( MeetingsTitle, "Meetings" ) & "</a>]  "
			if MustPrintFooterButton( IncludeStories, "Stories", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "stories.asp", "Stories" ) & Chr(34) & ">" &  GetButtonTitle( StoriesTitle, "Stories" ) & "</a>]  "
			if MustPrintFooterButton( IncludeCalendar, "Calendar", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "calendar.asp", "Calendar" ) & Chr(34) & ">" &  GetButtonTitle( CalendarTitle, "Calendar" ) & "</a>]  "
			if MustPrintFooterButton( IncludeLinks, "Links", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "links.asp", "Links" ) & Chr(34) & ">" &  GetButtonTitle( LinksTitle, "Links" ) & "</a>]  "
			if MustPrintFooterButton( IncludeQuotes, "Quotes", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "Quotes.asp", "Quotes" ) & Chr(34) & ">" &  GetButtonTitle( QuotesTitle, "Quotes" ) & "</a>]  "
			if MustPrintFooterButton( IncludeGuestbook, "Guestbook", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "guestbook.asp", "Guestbook" ) & Chr(34) & ">" &  GetButtonTitle( GuestbookTitle, "Guestbook" ) & "</a>]  "
			if MustPrintFooterButton( IncludeForum, "Forum", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "forum.asp", "Forum" ) & Chr(34) & ">" &  GetButtonTitle( ForumTitle, "Forum" ) & "</a>]  "
			if MustPrintFooterButton( IncludePhotos, "Photos", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "photos.asp", "Photos" ) & Chr(34) & ">" &  GetButtonTitle( PhotosTitle, "Photos" ) & "</a>]  "
			if MustPrintFooterButton( IncludeVoting, "Voting", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "voting.asp", "Voting" ) & Chr(34) & ">" &  GetButtonTitle( VotingTitle, "Voting" ) & "</a>]  "
			if MustPrintFooterButton( IncludeQuizzes, "Quizzes", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "quizzes.asp", "Quizzes" ) & Chr(34) & ">" &  GetButtonTitle( QuizzesTitle, "Quizzes" ) & "</a>]  "
			if MustPrintFooterButton( IncludeMedia, "Media", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "media.asp", "Media" ) & Chr(34) & ">" &  GetButtonTitle( MediaTitle, "Media" ) & "</a>]  "
			if MustPrintFooterButton( IncludeNewsletter, "Newsletter", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "Newsletter.asp", "Newsletter" ) & Chr(34) & ">" & GetButtonTitle( NewsletterTitle, "Newsletter" ) & "</a>]  "
			if MustPrintFooterButton( IncludeStore + AllowStore - 1, "Store", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "store.asp", "Store" ) & Chr(34) & ">" &  GetButtonTitle( StoreTitle, "Store" ) & "</a>]  "

			'Custom sub for the buttons
			if MustPrintFooterButton( 1, "Custom", i, blOverFlow ) then PrintCustomFooter

			if MustPrintFooterButton( IncludeStats, "Stats", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "stats.asp", "Stats" ) & Chr(34) & ">" & GetButtonTitle( StatsTitle, "Stats" ) & "</a>]  "
			if MustPrintFooterButton( 1, "Members", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "members.asp", "Members" ) & Chr(34) & ">" & GetButtonTitle( MembersTitle, "Members" ) & "</a>]  "

			if MustPrintFooterButton( 1, "Search", i, blOverFlow ) then FooterFile.Write "[<a href=" & Chr(34) & NonSecurePath & GetButtonLink( "search.asp", "Search" ) & Chr(34) & ">" & GetButtonTitle( "Search", "Search" ) & "</a>]  "

		next


		FooterFile.Write "</td></tr></table>"

	end if

	FooterFile.Write "<table border=0 width='100%'><tr><td align=" & rsLook("FooterAlignment") & ">Copyright &copy; 2001 <a href='http://www.GroupLoop.com'>www.GroupLoop.com</a>. All rights reserved.</td></tr></table>"
End Sub


HeaderFile.Close
FooterFile.Close
Set HeaderFile = Nothing
Set FooterFile = Nothing

Set FileSystem = Nothing

Set rsLook = Nothing
Set rsPages = Nothing

if Request("Source") <> "" and Request("Source") <> "No" then
	Redirect Request("Source")
elseif not (Request("Source") = "No" or strSource = "No") then
	Redirect "index.asp"
end if



Sub PrintButtons( OutFile, strSeparator, strAlignment, blPrintHeaderFooter )
	blImageExists = False

	if blPrintHeaderFooter then PrintImage blImageExists, strAlignment&"MenuTop", strAlignment&"MenuTop", OutFile, "", ""

	blFirst = (not blImageExists)

	blOverflow = false

	for i = 1 to intNumButtons + 1
		if i = intNumButtons + 1 then blOverflow = true

		if MustPrintButtonWithAlignment( strAlignment, 1, "Home", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Home", "Home", "index.asp?Action=Old", OutFile, strSeparator, ""

		if blParentSiteExists then
			if MustPrintButtonWithAlignment( strAlignment,1, "Parent"&ParentSites(0, 0), i, blOverFlow ) then
				PrintParentButton ParentSites(0, 0), strAlignment, blFirst, OutFile, strSeparator
			end if
		end if

		if blChildSiteExists then
			for p = 0 to intMaxChildSites - 1
				if MustPrintButtonWithAlignment( strAlignment,1, "Child"&ChildSites(p, 0), i, blOverFlow ) then
					PrintChildButton ChildSites(p, 0), strAlignment, blFirst, OutFile, strSeparator
				end if
			next
		end if

		for p = 0 to intMaxPages - 1
			if MustPrintButtonWithAlignment( strAlignment, 1, "InfoPage"&InfoPages(p, 0), i, blOverFlow ) then
				PrintButton strAlignment, blFirst, "InfoPage"&InfoPages(p, 0), InfoPages(p, 1), "pages_read.asp?ID=" & InfoPages(p, 0), OutFile, strSeparator, ""
			end if
		next

		if ButtonArray(i-1, 5) = 1 then
			if MustPrintButtonWithAlignment( strAlignment, 1, ButtonArray(i-1, 0), i, blOverFlow ) then PrintButton strAlignment, blFirst, ButtonArray(i-1, 0), ButtonArray(i-1, 3), ButtonArray(i-1, 4), OutFile, strSeparator, ""

		end if

		if MustPrintButtonWithAlignment( strAlignment, IncludeAnnouncements, "Announcements", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Announcements", AnnouncementsTitle, "announcements.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeMeetings, "Meetings", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Meetings", MeetingsTitle, "meetings.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeStories, "Stories", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Stories", StoriesTitle, "stories.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeCalendar, "Calendar", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Calendar", CalendarTitle, "calendar.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeLinks, "Links", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Links", LinksTitle, "links.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeQuotes, "Quotes", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Quotes", QuotesTitle, "quotes.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeGuestbook, "Guestbook", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Guestbook", GuestbookTitle, "guestbook.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeForum, "Forum", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Forum", ForumTitle, "forum.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludePhotos, "Photos", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Photos", PhotosTitle, "photos.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeVoting, "Voting", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Voting", VotingTitle, "voting.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeQuizzes, "Quizzes", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Quizzes", QuizzesTitle, "quizzes.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeMedia, "Media", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Media", MediaTitle, "media.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeNewsletter, "Newsletter", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Newsletter", NewsletterTitle, "newsletter.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, IncludeStore + AllowStore - 1, "Store", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Store", StoreTitle, "store.asp", OutFile, strSeparator, ""

		'Custom sub for the buttons
		if MustPrintButtonWithAlignment( strAlignment, 1, "Custom", i, blOverFlow ) then PrintCustomMenu

		if MustPrintButtonWithAlignment( strAlignment, IncludeStats, "Stats", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Stats", StatsTitle, "stats.asp", OutFile, strSeparator, ""
		if MustPrintButtonWithAlignment( strAlignment, 1, "Members", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Members", MembersTitle, "members.asp", OutFile, strSeparator, ""

		if MustPrintButtonWithAlignment( strAlignment, 1, "Search", i, blOverFlow ) then PrintButton strAlignment, blFirst, "Search", "Search", "search.asp", OutFile, strSeparator, ""

	next

	if blPrintHeaderFooter then PrintImage blImageExists, strAlignment&"MenuBottom", strAlignment&"MenuBottom", OutFile, strSeparator, ""

End Sub


'This is kinda like the mustprintbutton, but includes the alignment as well
Function MustPrintButtonWithAlignment( strAlignment, intInclude, strButton, i, blOverFlow )
	blMust = ( InStr(ButtonArray(i-1, 2), "Menu") and (ButtonArray(i-1, 1) = strAlignment) and MustPrintButtonNew( intInclude, strButton, i, blOverFlow ) )

	MustPrintButtonWithAlignment = blMust

End Function


'This is kinda like the mustprintbutton, but checks if it needs printed in the footer
Function MustPrintFooterButton( intInclude, strButton, i, blOverFlow )
	blMust = ( InStr(ButtonArray(i-1, 2), "Footer") and MustPrintButtonNew( intInclude, strButton, i, blOverFlow ) )

	MustPrintFooterButton = blMust

End Function


Sub PrintButton( strMenuAlignment, blFirst, strImage, strTitle, strLink, WriteFile, strHeader, strAltPath )

	strButtonName = strImage
	strTitle = GetButtonTitle( strTitle, strButtonName )

	strButtonHeader = rsLook(strMenuAlignment & "MenuButtonHeaderText")
	strButtonFooter = rsLook(strMenuAlignment & "MenuButtonFooterText")

	'If it's the first button, do not print the header
	if not blFirst then WriteFile.Write strHeader

	blFirst = false

	strName = strImage & "Image"
	strOverName = strImage & "RolloverImage"
	strHeaderName = strImage & "HeaderImage"
	strFooterName = strImage & "FooterImage"


	strExt = ""
	strOverExt = ""
	strHeaderExt = ""
	strFooterExt = ""

	blImage = ImageExists( strName, strExt )
	blRollover = ImageExists( strOverName, strOverExt )
	blHeader = ImageExists( strHeaderName, strHeaderExt )
	blFooter = ImageExists( strFooterName, strFooterExt )

	strPath = NonSecurePath
	if strAltPath <> "" then strPath = strAltPath

	WriteFile.Write strButtonHeader

	if blHeader then WriteFile.Write "<img src='images/"& strHeaderName & "." & strHeaderExt & "' border=0>" & strHeader


	if blImage and not blRollover then
		WriteFile.Write "<a href=" & Chr(34) & strPath & strLink & Chr(34) & "><img src='images/"& strName & "." & strExt & "' alt='" & strTitle & "' border=0></a>"
	elseif blImage and blRollover then
		WriteFile.Write "<a href=" & Chr(34) & strPath & strLink & Chr(34) & " onMouseOver=" & Chr(34) & "document.images['" & strTitle & "'].src='images/"& strOverName & "." & strOverExt & "'" & Chr(34) & " onMouseOut=" & Chr(34) & "document.images['" & strTitle & "'].src='images/" & strName & "." & strExt & "'" & Chr(34) & "><img src='images/" & strName & "." & strExt & "' alt='" & strTitle & "' border=0 name='" & strTitle & "'></a>"
	else
		WriteFile.Write "<a class='MenuLink' href=" & Chr(34) & strPath & strLink & Chr(34) & "><span class='" & strMenuAlignment & "Menu'>" & strTitle & "</span></a>"
	end if

	WriteFile.Write strButtonFooter


	if blFooter then WriteFile.Write strHeader & "<img src='images/"& strFooterName & "." & strFooterExt & "' border=0>"


End Sub


Sub PrintImage( blExists, strImage, strTitle, WriteFile, strHeader, strFooter )
	strName = strImage & "Image"
	strOverName = strImage & "RolloverImage"

	strExt = ""
	strOverExt = ""

	blImage = ImageExists( strName, strExt )
	blRollover = ImageExists( strOverName, strOverExt )

	blExists = blImage or blRollover

	if blImage and not blRollover then
		WriteFile.Write  strHeader & "<img src='images/"& strName & "." & strExt & "' border=0>" & strFooter
	elseif blImage and blRollover then
		WriteFile.Write strHeader & "<a href=" & Chr(34) & "#" & Chr(34) & " onMouseOver=" & Chr(34) & "document.images['" & strTitle & "'].src='images/"& strOverName & "." & strOverExt & "'" & Chr(34) & " onMouseOut=" & Chr(34) & "document.images['" & strTitle & "'].src='images/" & strName & "." & strExt & "'" & Chr(34) & "><img src='images/" & strName & "." & strExt & "' border=0 name='" & strTitle & "'></a>" & strFooter
	end if
End Sub



Sub PrintChildButton( intChildID, strAlignment, blFirst, WriteFile, strFooter )
	GetChild intChildID, strShortTitle, strSubDirectory

	'Erase the first directory, so we don't get www.grouploop.com/parent/parent/child
	intPos = InstrRev(strSubDirectory, "/", (Len(strSubDirectory) - 1) )
	strSubDirectory = Right( strSubDirectory, Len(strSubDirectory) - intPos )

	'if we are a parent site, just tack on the extra directory.  if we are a child, earse the subdir and tack on this one
	if Version = "Parent" then
		strAlt = NonSecurePath & strSubDirectory
	else
		intLastPos = InstrRev(NonSecurePath, "/", (Len(NonSecurePath) - 1) )
		strAlt = Left( NonSecurePath, intLastPos ) & strSubDirectory
	end if

	PrintButton strAlignment, blFirst, "Child" & intChildID, strShortTitle, "", WriteFile, strFooter, strAlt
End Sub


Sub PrintParentButton( intParentID, strAlignment, blFirst, WriteFile, strFooter )
	GetParent intParentID, strShortTitle, strSubDirectory

	'Erase our sub-directory
	intLastPos = InstrRev(NonSecurePath, "/", (Len(NonSecurePath) - 1) )
	strAlt = Left( NonSecurePath, intLastPos )

	'If the parent is also a child, then add its directory to it
	if Instr( strSubDirectory, "/" ) then
		intLastPos = InstrRev(strSubDirectory, "/", (Len(strSubDirectory) - 1) )
		strAlt = strAlt & Right( strSubDirectory, (Len(strSubDirectory) - intLastPos) ) & "/"
		Response.Write strAlt & "<br>"
	end if

	PrintButton strAlignment, blFirst, "Parent" & intParentID, strShortTitle & " Home", "", WriteFile, strFooter, strAlt
End Sub

Sub GetParent( intParentID, strShortTitle, strSubDirectory )
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
		strShortTitle = .Parameters("@ShortTitle")
		strSubDirectory = .Parameters("@SubDirectory")
	End With
	Set cmdTemp = Nothing
End Sub


Sub GetChild( intChildID, strShortTitle, strSubDirectory )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetChildSiteInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ShortTitle", adVarWChar, adParamOutput, 100 )
		.Parameters.Append .CreateParameter ("@SubDirectory", adVarWChar, adParamOutput, 100 )
		.Parameters("@CustomerID") = intChildID
		.Execute , , adExecuteNoRecords
		strShortTitle = .Parameters("@ShortTitle")
		strSubDirectory = .Parameters("@SubDirectory")
	End With
	Set cmdTemp = Nothing

	GetChildID = intChildID
End Sub



'This function simply returns the width output for a menu if there is one
Function GetMenuWidth( strMenuAlignment )
	if IsNull(rsLook(strMenuAlignment & "MenuWidth")) then
		strWidth = ""
	elseif rsLook(strMenuAlignment & "MenuWidth") = "" then
		strWidth = ""
	else
		strWidth = " width = " & Chr(34) & rsLook(strMenuAlignment & "MenuWidth") & Chr(34) & " "
	end if

	GetMenuWidth = strWidth
End Function

'This returns either the regular title, or a custom one they entered
Function GetButtonTitle( strDefaultTitle, strButtonName )

	strTOut = strDefaultTitle

	for x = 0 to intNumButtons - 1
		if ButtonArray(x, 0) = strButtonName then
			if not IsNull(ButtonArray(x, 3)) then
				'They have a custom title
				if ButtonArray(x, 3) <> "" then strTOut = ButtonArray(x, 3)
			end if
		end if
	next

	GetButtonTitle = strTOut

End Function

'This returns either the regular Link, or a custom one they entered
Function GetButtonLink( strDefaultLink, strButtonName )

	strTOut = strDefaultLink

	for x = 0 to intNumButtons - 1
		if ButtonArray(x, 0) = strButtonName then
			if not IsNull(ButtonArray(x, 3)) then
				'They have a custom Link
				if ButtonArray(x, 4) <> "" then strTOut = ButtonArray(x, 4)
			end if
		end if
	next

	GetButtonLink = strTOut

End Function


'This gives the separator between buttons for a specific menu. It can be text, image, or both
Function GetSeparator( strSector )
	if ImageExists( strSector & "MenuSeparatorImage", strExt ) then
		strSeparator = rsLook(strSector & "MenuSeparatorHeaderText") & "<img src='images/" & strSector & "MenuSeparatorImage." & strExt & "' border='0'>" & rsLook(strSector & "MenuSeparatorFooterText")
	else
		strSeparator = rsLook(strSector & "MenuSeparator")
	end if
	GetSeparator = strSeparator
End Function


'Puts the database info into an array for faster searching
Sub FillArrays
	blButtons = SectorHasButtons("Top") or SectorHasButtons("Left") or SectorHasButtons("Right")

	Set rsItems = Server.CreateObject("ADODB.Recordset")

	if blButtons then
		Query = "SELECT * FROM MenuButtons WHERE CustomerID = " & CustomerID & " ORDER BY Position"
	else
		Query = "SELECT * FROM MenuButtons WHERE CustomerID = -1 ORDER BY Position"
	end if
	rsItems.CacheSize = 50
	rsItems.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsItems.EOF then
		intNumButtons = 0

	else
		intNumButtons = rsItems.RecordCount
		for a = 0 to intNumButtons - 1
			ButtonArray(a, 0) = rsItems("Name")
			
//			if not Instr(rsItems("Name"), "InfoPage") then blButtons = False

			ButtonArray(a, 1) = rsItems("Align")
			ButtonArray(a, 2) = rsItems("Show")
			ButtonArray(a, 3) = rsItems("CustomLabel")
			ButtonArray(a, 4) = rsItems("CustomLink")
			ButtonArray(a, 5) = rsItems("Custom")
			rsItems.MoveNext
		next
	end if


	rsItems.Close

	Query = "SELECT ID, Title FROM InfoPages WHERE Title <> 'Home Page' AND CustomerID = " & CustomerID & " AND (ShowButton = 1) ORDER BY Title"

	rsItems.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	intMaxPages = 0
	if not rsItems.EOF then
		Set PID = rsItems("ID")
		Set PTitle = rsItems("Title")
		intMaxPages = rsItems.RecordCount


		for a = 0 to intMaxPages - 1
			InfoPages(a, 0) = rsItems("ID")
			InfoPages(a, 1) = rsItems("Title")
			rsItems.MoveNext
		next
	end if
	rsItems.Close

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
			intMaxChildSites = rsItems.RecordCount
			for a = 0 to intMaxChildSites - 1
				ChildSites(a, 0) = rsItems("ID")
				ChildSites(a, 1) = rsItems("Title")
				ChildSites(a, 2) = rsItems("Subdirectory")
				rsItems.MoveNext
			next
		end if
		rsItems.Close

		Set cmdTemp = Nothing
	end if


	if blParentSiteExists then
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

			ParentSites(0, 0) = .Parameters("@ParentID")
			ParentSites(0, 1) = .Parameters("@ShortTitle")
			ParentSites(0, 2) = .Parameters("@SubDirectory")


		End With
		Set cmdTemp = Nothing
	end if

	Set rsItems = Nothing

End Sub


Function MustPrintButtonNew( intInclude, strButton, i, blOverFlow )
	MustPrintButtonNew = intInclude > 0 and ( ButtonArray(i-1, 0) = strButton or (blOverFlow and not InNumButtonArray( strButton )) )
End Function


Function InNumButtonArray( strButton )
	for x = 0 to intNumButtons - 1
		if ButtonArray(x, 0) = strButton then
			InNumButtonArray = true
			exit function
		end if
	next

	InNumButtonArray = false
End Function

%>
