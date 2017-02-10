<!-- #include file="admin_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Basic Graphics!!!</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'update info
Set upl = Server.CreateObject("SoftArtisans.FileUp")
strPath = GetPath ("images")
upl.Path = strPath

Set FileSystem = CreateObject("Scripting.FileSystemObject")

if not LoggedAdmin and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedAdmin then
	Set upl = Nothing
	Set FileSystem = Nothing
	Redirect("members.asp?Source=admin_images_edit.asp")
end if


strError = ""

SetImage "BackgroundImage"

SetImage "TitleImage"
SetImage "TitleRolloverImage"
SetImage "TitleMenuBackgroundImage"
SetImage "TitleTopImage"
SetImage "TitleTopRolloverImage"
SetImage "TitleBottomImage"
SetImage "TitleBottomRolloverImage"

SetImage "TopMenuBackgroundImage"
SetImage "TopMenuTopImage"
SetImage "TopMenuTopRolloverImage"
SetImage "TopMenuBottomImage"
SetImage "TopMenuBottomRolloverImage"
SetImage "TopMenuSeparatorImage"

SetImage "LeftMenuBackgroundImage"
SetImage "LeftMenuTopImage"
SetImage "LeftMenuTopRolloverImage"
SetImage "LeftMenuBottomImage"
SetImage "LeftMenuBottomRolloverImage"
SetImage "LeftMenuSeparatorImage"

SetImage "RightMenuBackgroundImage"
SetImage "RightMenuTopImage"
SetImage "RightMenuTopRolloverImage"
SetImage "RightMenuBottomImage"
SetImage "RightMenuBottomRolloverImage"
SetImage "RightMenuSeparatorImage"

SetImage "BodyMenuBackgroundImage"


SetImage "HomeImage"
SetImage "HomeRolloverImage"
SetImage "HomeHeaderImage"
SetImage "HomeFooterImage"

SetImage "BulletImage"
if CBool( IncludeAnnouncements ) then SetImage "AnnouncementsImage"
if CBool( IncludeAnnouncements ) then SetImage "AnnouncementsRolloverImage"
if CBool( IncludeAnnouncements ) then SetImage "AnnouncementsHeaderImage"
if CBool( IncludeAnnouncements ) then SetImage "AnnouncementsFooterImage"

if CBool( IncludeMeetings ) then SetImage "MeetingsImage"
if CBool( IncludeMeetings ) then SetImage "MeetingsRolloverImage"
if CBool( IncludeMeetings ) then SetImage "MeetingsHeaderImage"
if CBool( IncludeMeetings ) then SetImage "MeetingsFooterImage"

if CBool( IncludeStories ) then SetImage "StoriesImage"
if CBool( IncludeStories ) then SetImage "StoriesRolloverImage"
if CBool( IncludeStories ) then SetImage "StoriesHeaderImage"
if CBool( IncludeStories ) then SetImage "StoriesFooterImage"

if CBool( IncludeCalendar ) then SetImage "CalendarImage"
if CBool( IncludeCalendar ) then SetImage "CalendarRolloverImage"
if CBool( IncludeCalendar ) then SetImage "CalendarHeaderImage"
if CBool( IncludeCalendar ) then SetImage "CalendarFooterImage"

if CBool( IncludeCalendar ) then SetImage "LastMonthImage"
if CBool( IncludeCalendar ) then SetImage "LastMonthRolloverImage"
if CBool( IncludeCalendar ) then SetImage "NextMonthImage"
if CBool( IncludeCalendar ) then SetImage "NextMonthRolloverImage"

if CBool( IncludeQuotes ) then SetImage "QuotesImage"
if CBool( IncludeQuotes ) then SetImage "QuotesRolloverImage"
if CBool( IncludeQuotes ) then SetImage "QuotesHeaderImage"
if CBool( IncludeQuotes ) then SetImage "QuotesFooterImage"

if CBool( IncludeLinks ) then SetImage "LinksImage"
if CBool( IncludeLinks ) then SetImage "LinksRolloverImage"
if CBool( IncludeLinks ) then SetImage "LinksHeaderImage"
if CBool( IncludeLinks ) then SetImage "LinksFooterImage"

if CBool( IncludeGuestbook ) then SetImage "GuestbookImage"
if CBool( IncludeGuestbook ) then SetImage "GuestbookRolloverImage"
if CBool( IncludeGuestbook ) then SetImage "GuestbookHeaderImage"
if CBool( IncludeGuestbook ) then SetImage "GuestbookFooterImage"

if CBool( IncludeForum ) then
	SetImage "ForumImage"
	SetImage "ForumRolloverImage"
	SetImage "ForumHeaderImage"
	SetImage "ForumFooterImage"
end if

if CBool( IncludeForum ) then SetImage "ForumPlusImage"
if CBool( IncludeForum ) then SetImage "ForumMinusImage"
if CBool( IncludePhotos ) then SetImage "PhotosImage"
if CBool( IncludePhotos ) then SetImage "PhotosRolloverImage"
if CBool( IncludePhotos ) then SetImage "PhotosHeaderImage"
if CBool( IncludePhotos ) then SetImage "PhotosFooterImage"

if CBool( IncludeVoting ) then SetImage "VotingImage"
if CBool( IncludeVoting ) then SetImage "VotingRolloverImage"
if CBool( IncludeVoting ) then SetImage "VotingHeaderImage"
if CBool( IncludeVoting ) then SetImage "VotingFooterImage"

if CBool( IncludeQuizzes ) then SetImage "QuizzesImage"
if CBool( IncludeQuizzes ) then SetImage "QuizzesRolloverImage"
if CBool( IncludeQuizzes ) then SetImage "QuizzesHeaderImage"
if CBool( IncludeQuizzes ) then SetImage "QuizzesFooterImage"

if CBool( IncludeMedia ) then SetImage "MediaImage"
if CBool( IncludeMedia ) then SetImage "MediaRolloverImage"
if CBool( IncludeMedia ) then SetImage "MediaHeaderImage"
if CBool( IncludeMedia ) then SetImage "MediaFooterImage"

if CBool( IncludeNewsletter ) then SetImage "NewsletterImage"
if CBool( IncludeNewsletter ) then SetImage "NewsletterRolloverImage"
if CBool( IncludeNewsletter ) then SetImage "NewsletterHeaderImage"
if CBool( IncludeNewsletter ) then SetImage "NewsletterFooterImage"

if CBool( AllowStore ) AND CBool( IncludeStore ) then SetImage "StoreImage"
if CBool( AllowStore ) AND CBool( IncludeStore ) then SetImage "StoreRolloverImage"
if CBool( AllowStore ) AND CBool( IncludeStore ) then SetImage "StoreHeaderImage"
if CBool( AllowStore ) AND CBool( IncludeStore ) then SetImage "StoreFooterImage"


if CBool( IncludeStats ) then SetImage "StatsImage"
if CBool( IncludeStats ) then SetImage "StatsRolloverImage"
if CBool( IncludeStats ) then SetImage "StatsHeaderImage"
if CBool( IncludeStats ) then SetImage "StatsFooterImage"

SetImage "MembersImage"
SetImage "MembersRolloverImage"
SetImage "MembersHeaderImage"
SetImage "MembersFooterImage"

SetImage "SearchImage"
SetImage "SearchRolloverImage"
SetImage "SearchHeaderImage"
SetImage "SearchFooterImage"


if CBool( TellNew ) then SetImage "NewImage"


Set rsPages = Server.CreateObject("ADODB.Recordset")

if ChildSiteExists() then
	Query = "SELECT ID FROM Customers WHERE ParentID = " & CustomerID
	rsPages.CacheSize = 20
	rsPages.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	if not rsPages.EOF then	Set ChildID = rsPages("ID")
	do until rsPages.EOF
		SetImage "Child" & ChildID & "Image"
		SetImage "Child" & ChildID & "RolloverImage"
		SetImage "Child" & ChildID & "HeaderImage"
		SetImage "Child" & ChildID & "FooterImage"
		rsPages.MoveNext
	loop
	rsPages.Close
end if

if ParentSiteExists() then
	GetParent intParentID, strShortTitle, strSubDirectory
	SetImage "Parent" & intParentID & "Image"
	SetImage "Parent" & intParentID & "RolloverImage"
	SetImage "Parent" & intParentID & "HeaderImage"
	SetImage "Parent" & intParentID & "FooterImage"
end if


Query = "SELECT ID FROM InfoPages WHERE Title <> 'Home Page' AND CustomerID = " & CustomerID & " AND ShowButton = 1"
rsPages.CacheSize = 20
rsPages.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
if not rsPages.EOF then

	Set ID = rsPages("ID")

	do until rsPages.EOF
		SetImage "InfoPage" & ID & "Image"
		SetImage "InfoPage" & ID & "RolloverImage"
		SetImage "InfoPage" & ID & "HeaderImage"
		SetImage "InfoPage" & ID & "FooterImage"
		rsPages.MoveNext
	loop
end if
rsPages.Close
Set rsPages = Nothing

Set upl = Nothing

'Reset the path so it doesn't interfere with the constants file
strPath = ""
%>
	<!-- #include file="write_constants.asp" -->
<%
if strError = "" then
	Redirect("write_header_footer.asp?Source=admin_images_edit.asp?Submit=Changed")
else
	Redirect("message.asp?Source=admin_images_edit.asp&Message=" & Server.URLEncode( strError ))
end if

Set FileSystem = Nothing

'-------------------------------------------------------------
'Take in a field name and either upload its image, delete it, or do nothing
'-------------------------------------------------------------
Sub SetImage( strField )
	'Get the image
	if IsObject(upl.Form("Up"&strField)) then
		if not upl.Form("Up"&strField).IsEmpty then
			'Delete the old one if there is one...
			if FileSystem.FileExists( strPath & strField & ".jpg" ) then FileSystem.DeleteFile( strPath & strField & ".jpg" )
			if FileSystem.FileExists( strPath & strField & ".gif" ) then FileSystem.DeleteFile( strPath & strField & ".gif" )
			if FileSystem.FileExists( strPath & strField & ".bmp" ) then FileSystem.DeleteFile( strPath & strField & ".bmp" )

			'--- Retrieve the file's content type and assign it to a variable
			FTYPE = upl.Form("Up"&strField).ContentType	

			'--- Restrict the file types saved using a Select condition
			if FTYPE = "image/gif" then
				upl.Form("Up"&strField).SaveAs strField&".gif"
			elseif FTYPE = "image/pjpeg" or FTYPE = "image/jpeg" then
				upl.Form("Up"&strField).SaveAs strField&".jpg"
			elseif FTYPE = "image/bmp" then
				upl.Form("Up"&strField).SaveAs strField&".bmp"
			else
				upl.Form("Up"&strField).delete
				strError = strError & "You can only upload gif, jpeg, and bitmap (bmp) images.  For " & strField & " you uploaded an invalid type of image.<br>"
			end if

		'Erase the image
		elseif upl.Form(strField) = "0" then
			if FileSystem.FileExists( strPath & strField & ".jpg" ) then FileSystem.DeleteFile( strPath & strField & ".jpg" )
			if FileSystem.FileExists( strPath & strField & ".gif" ) then FileSystem.DeleteFile( strPath & strField & ".gif" )
			if FileSystem.FileExists( strPath & strField & ".bmp" ) then FileSystem.DeleteFile( strPath & strField & ".bmp" )
		end if
	end if
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
'------------------------End Code-----------------------------
%>
