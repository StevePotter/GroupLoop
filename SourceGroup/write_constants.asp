<%
'This script takes all the information from the configuration table and puts into a file defining 
'constants for each field.  Avoids an unnecessary recordset for every page

'-----------------------Create Files----------------------------
if strPath = "" then strPath = GetPath("")
strImagePath = strPath & "images/"

if not IsObject(FileSystem) then
	blFileSystemExisted = false
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
else
	blFileSystemExisted = true
end if

Set ConstFile = FileSystem.CreateTextFile(strPath & "constantstemp.inc")

DblQuote = Chr(34)

if CustomerID = "" then CustomerID = intCustomerID


'Open up the configuration recordset
Query = "SELECT * FROM Configuration WHERE CustomerID = " & CustomerID

if not IsObject(rsConfig) then
	blrsConfigExisted = false
else
	blrsConfigExisted = true
end if

Set rsConfig = Server.CreateObject("ADODB.Recordset")

rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect


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
	strVersion = .Parameters("@Version")
	intDomain = .Parameters("@UseDomain")
	strDomainName = .Parameters("@DomainName")
	intMasterID = .Parameters("@MasterID")

	'If this is a sub-site, get the master domain (if there is one...)
	if intMasterID > 0 then
		.Parameters("@CustomerID") = intMasterID
		.Execute , , adExecuteNoRecords
		intDomain = .Parameters("@UseDomain")
		strDomainName = .Parameters("@DomainName")
	end if

End With

Set Command = Nothing

'-----------------------Start Writing Files----------------------------

'Create our opening delimiter
ConstFile.Write "<"
ConstFile.WriteLine "%"
ConstFile.WriteBlankLines 2

ConstFile.WriteLine "'This file contains the constants for this site's configuration table (cust #" & CustomerID

ConstFile.WriteBlankLines 2

ConstFile.WriteLine "Const Version = " & DblQuote & strVersion & DblQuote
ConstFile.WriteLine "Const SiteDate = " & DblQuote & FormatDateTime(rsConfig("Date"), 2) & DblQuote
if CustomerID = "" then
	ConstFile.WriteLine "Const CustomerID = " & CustomerID
else
	ConstFile.WriteLine "Const CustomerID = " & CustomerID
end if

'Get all the images
strExt = ""
ConstFile.WriteLine "Const ForumPlusImage = " & ImageExists( "ForumPlusImage", strExt )
ConstFile.WriteLine "Const ForumPlusImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const ForumMinusImage = " & ImageExists( "ForumMinusImage", strExt )
ConstFile.WriteLine "Const ForumMinusImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const LastMonthImage = " & ImageExists( "LastMonthImage", strExt )
ConstFile.WriteLine "Const LastMonthImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const LastMonthRolloverImage = " & ImageExists( "LastMonthRolloverImage", strExt )
ConstFile.WriteLine "Const LastMonthRolloverImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const NextMonthImage = " & ImageExists( "NextMonthImage", strExt )
ConstFile.WriteLine "Const NextMonthImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const NextMonthRolloverImage = " & ImageExists( "NextMonthRolloverImage", strExt )
ConstFile.WriteLine "Const NextMonthRolloverImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const NewImage = " & ImageExists("NewImage", strExt )
ConstFile.WriteLine "Const NewImageExt = " & DblQuote & strExt & DblQuote
ConstFile.WriteLine "Const BulletImage = " & ImageExists("BulletImage", strExt )
ConstFile.WriteLine "Const BulletImageExt = " & DblQuote & strExt & DblQuote

ConstFile.WriteLine "Const CellSpacing = " & rsConfig("CellSpacing")
ConstFile.WriteLine "Const CellPadding = " & rsConfig("CellPadding")
ConstFile.WriteLine "Const Border = " & rsConfig("Border")
ConstFile.WriteLine "Const HeadingAlignment = " & DblQuote & rsConfig("HeadingAlignment") & DblQuote
ConstFile.WriteLine "Const PageSize = " & rsConfig("PageSize")

ConstFile.WriteLine "Const IncludeNewsletter = " & rsConfig("IncludeNewsletter")
ConstFile.WriteLine "Const IncludeStories = " & rsConfig("IncludeStories")
ConstFile.WriteLine "Const IncludeAnnouncements = " & rsConfig("IncludeAnnouncements")
ConstFile.WriteLine "Const IncludeCalendar = " & rsConfig("IncludeCalendar")
ConstFile.WriteLine "Const IncludeQuizzes = " & rsConfig("IncludeQuizzes")
ConstFile.WriteLine "Const IncludeVoting = " & rsConfig("IncludeVoting")
ConstFile.WriteLine "Const IncludePhotos = " & rsConfig("IncludePhotos")
ConstFile.WriteLine "Const IncludePhotoCaptions = " & rsConfig("IncludePhotoCaptions")
ConstFile.WriteLine "Const IncludeGuestbook = " & rsConfig("IncludeGuestbook")
ConstFile.WriteLine "Const IncludeForum = " & rsConfig("IncludeForum")
ConstFile.WriteLine "Const IncludeStats = " & rsConfig("IncludeStats")
ConstFile.WriteLine "Const IncludeLinks = " & rsConfig("IncludeLinks")
ConstFile.WriteLine "Const IncludeAdditions = " & rsConfig("IncludeAdditions")
ConstFile.WriteLine "Const IncludeQuotes = " & rsConfig("IncludeQuotes")
ConstFile.WriteLine "Const IncludeMedia = " & rsConfig("IncludeMedia")
ConstFile.WriteLine "Const IncludeMeetings = " & rsConfig("IncludeMeetings")
ConstFile.WriteLine "Const IncludeStore = " & rsConfig("IncludeStore")
ConstFile.WriteLine "Const IncludeAuthor = " & rsConfig("IncludeAuthor")
ConstFile.WriteLine "Const IncludeDate = " & rsConfig("IncludeDate")
ConstFile.WriteLine "Const IncludeAddButtons = " & rsConfig("IncludeAddButtons")
ConstFile.WriteLine "Const IncludeCommittees = " & rsConfig("IncludeCommittees")
ConstFile.WriteLine "Const IncludeEditSectionPropButtons = " & rsConfig("IncludeEditSectionPropButtons")


ConstFile.WriteLine "Const SectionViewAnnouncements = " & DblQuote & rsConfig("SectionViewAnnouncements") & DblQuote
ConstFile.WriteLine "Const SectionViewCalendar = " & DblQuote & rsConfig("SectionViewCalendar") & DblQuote
ConstFile.WriteLine "Const SectionViewNewsletter = " & DblQuote & rsConfig("SectionViewNewsletter") & DblQuote
ConstFile.WriteLine "Const SectionViewAdditions = " & DblQuote & rsConfig("SectionViewAdditions") & DblQuote
ConstFile.WriteLine "Const SectionViewQuizzes = " & DblQuote & rsConfig("SectionViewQuizzes") & DblQuote
ConstFile.WriteLine "Const SectionViewVoting = " & DblQuote & rsConfig("SectionViewVoting") & DblQuote
ConstFile.WriteLine "Const SectionViewPhotos = " & DblQuote & rsConfig("SectionViewPhotos") & DblQuote
ConstFile.WriteLine "Const SectionViewPhotoCaptions = " & DblQuote & rsConfig("SectionViewPhotoCaptions") & DblQuote
ConstFile.WriteLine "Const SectionViewLinks = " & DblQuote & rsConfig("SectionViewLinks") & DblQuote
ConstFile.WriteLine "Const SectionViewNews = " & DblQuote & rsConfig("SectionViewNews") & DblQuote
ConstFile.WriteLine "Const SectionViewStats = " & DblQuote & rsConfig("SectionViewStats") & DblQuote
ConstFile.WriteLine "Const SectionViewQuotes = " & DblQuote & rsConfig("SectionViewQuotes") & DblQuote
ConstFile.WriteLine "Const SectionViewMedia = " & DblQuote & rsConfig("SectionViewMedia") & DblQuote
ConstFile.WriteLine "Const SectionViewMeetings = " & DblQuote & rsConfig("SectionViewMeetings") & DblQuote



ConstFile.WriteLine "Const AllowStore = " & rsConfig("AllowStore")
ConstFile.WriteLine "Const NewsletterMembers = " & rsConfig("NewsletterMembers")
ConstFile.WriteLine "Const QuizzesMembers = " & rsConfig("QuizzesMembers")
ConstFile.WriteLine "Const VotingMembers = " & rsConfig("VotingMembers")
ConstFile.WriteLine "Const PhotosMembers = " & rsConfig("PhotosMembers")
ConstFile.WriteLine "Const ForumMembers = " & rsConfig("ForumMembers")
ConstFile.WriteLine "Const InsertsMembers = " & rsConfig("InsertsMembers")
ConstFile.WriteLine "Const InfoPagesMembers = " & rsConfig("InfoPagesMembers")
ConstFile.WriteLine "Const MediaMembers = " & rsConfig("MediaMembers")
ConstFile.WriteLine "Const MeetingsMembers = " & rsConfig("MeetingsMembers")
ConstFile.WriteLine "Const CalendarMembers = " & rsConfig("CalendarMembers")
ConstFile.WriteLine "Const AnnouncementsMembers = " & rsConfig("AnnouncementsMembers")
ConstFile.WriteLine "Const StoriesMembers = " & rsConfig("StoriesMembers")
ConstFile.WriteLine "Const QuotesMembers = " & rsConfig("QuotesMembers")
ConstFile.WriteLine "Const LinksMembers = " & rsConfig("LinksMembers")

ConstFile.WriteLine "Const Title = " & DblQuote & rsConfig("Title") & DblQuote
ConstFile.WriteLine "Const NewsletterTitle = " & DblQuote & rsConfig("NewsletterTitle") & DblQuote
ConstFile.WriteLine "Const AdditionsTitle = " & DblQuote & rsConfig("AdditionsTitle") & DblQuote
ConstFile.WriteLine "Const AnnouncementsTitle = " & DblQuote & rsConfig("AnnouncementsTitle") & DblQuote
ConstFile.WriteLine "Const StoriesTitle = " & DblQuote & rsConfig("StoriesTitle") & DblQuote
ConstFile.WriteLine "Const CalendarTitle = " & DblQuote & rsConfig("CalendarTitle") & DblQuote
ConstFile.WriteLine "Const QuizzesTitle = " & DblQuote & rsConfig("QuizzesTitle") & DblQuote
ConstFile.WriteLine "Const VotingTitle = " & DblQuote & rsConfig("VotingTitle") & DblQuote
ConstFile.WriteLine "Const PhotosTitle = " & DblQuote & rsConfig("PhotosTitle") & DblQuote
ConstFile.WriteLine "Const PhotoCaptionsTitle = " & DblQuote & rsConfig("PhotoCaptionsTitle") & DblQuote
ConstFile.WriteLine "Const ForumTitle = " & DblQuote & rsConfig("ForumTitle") & DblQuote
ConstFile.WriteLine "Const MembersTitle = " & DblQuote & rsConfig("MembersTitle") & DblQuote
ConstFile.WriteLine "Const GuestbookTitle = " & DblQuote & rsConfig("GuestbookTitle") & DblQuote
ConstFile.WriteLine "Const LinksTitle = " & DblQuote & rsConfig("LinksTitle") & DblQuote
ConstFile.WriteLine "Const NewsTitle = " & DblQuote & rsConfig("NewsTitle") & DblQuote
ConstFile.WriteLine "Const AboutSiteTitle = " & DblQuote & rsConfig("AboutSiteTitle") & DblQuote
ConstFile.WriteLine "Const StatsTitle = " & DblQuote & rsConfig("StatsTitle") & DblQuote
ConstFile.WriteLine "Const QuotesTitle = " & DblQuote & rsConfig("QuotesTitle") & DblQuote
ConstFile.WriteLine "Const MediaTitle = " & DblQuote & rsConfig("MediaTitle") & DblQuote
ConstFile.WriteLine "Const MeetingsTitle = " & DblQuote & rsConfig("MeetingsTitle") & DblQuote
ConstFile.WriteLine "Const StoreTitle = " & DblQuote & rsConfig("StoreTitle") & DblQuote

ConstFile.WriteLine "Const UsernameLabel = " & DblQuote & rsConfig("UsernameLabel") & DblQuote

ConstFile.WriteLine "Const TellNew = " & rsConfig("TellNew")
ConstFile.WriteLine "Const AdditionsDaysOld = " & rsConfig("AdditionsDaysOld")
ConstFile.WriteLine "Const NewDaysOld = " & rsConfig("NewDaysOld")

ConstFile.WriteLine "Const RateAnnouncements = " & rsConfig("RateAnnouncements")
ConstFile.WriteLine "Const RateCalendar = " & rsConfig("RateCalendar")
ConstFile.WriteLine "Const RateStories = " & rsConfig("RateStories")
ConstFile.WriteLine "Const RateForum = " & rsConfig("RateForum")
ConstFile.WriteLine "Const RatePhotos = " & rsConfig("RatePhotos")
ConstFile.WriteLine "Const RatePhotoCaptions = " & rsConfig("RatePhotoCaptions")
ConstFile.WriteLine "Const RateQuizzes = " & rsConfig("RateQuizzes")
ConstFile.WriteLine "Const RateVoting = " & rsConfig("RateVoting")
ConstFile.WriteLine "Const RateLinks = " & rsConfig("RateLinks")
ConstFile.WriteLine "Const RateGuestbook = " & rsConfig("RateGuestbook")
ConstFile.WriteLine "Const RateQuotes = " & rsConfig("RateQuotes")
ConstFile.WriteLine "Const RateMedia = " & rsConfig("RateMedia")
ConstFile.WriteLine "Const RateMeetings = " & rsConfig("RateMeetings")
ConstFile.WriteLine "Const RateMembers = " & rsConfig("RateMembers")

ConstFile.WriteLine "Const ReviewAnnouncements = " & rsConfig("ReviewAnnouncements")
ConstFile.WriteLine "Const ReviewCalendar = " & rsConfig("ReviewCalendar")
ConstFile.WriteLine "Const ReviewStories = " & rsConfig("ReviewStories")
ConstFile.WriteLine "Const ReviewPhotoCaptions = " & rsConfig("ReviewPhotoCaptions")
ConstFile.WriteLine "Const ReviewQuizzes = " & rsConfig("ReviewQuizzes")
ConstFile.WriteLine "Const ReviewVoting = " & rsConfig("ReviewVoting")
ConstFile.WriteLine "Const ReviewLinks = " & rsConfig("ReviewLinks")
ConstFile.WriteLine "Const ReviewGuestbook = " & rsConfig("ReviewGuestbook")
ConstFile.WriteLine "Const ReviewQuotes = " & rsConfig("ReviewQuotes")
ConstFile.WriteLine "Const ReviewMedia = " & rsConfig("ReviewMedia")
ConstFile.WriteLine "Const ReviewMembers = " & rsConfig("ReviewMembers")
ConstFile.WriteLine "Const ReviewMeetings = " & rsConfig("ReviewMeetings")

ConstFile.WriteLine "Const VotingBarColor = " & DblQuote & rsConfig("VotingBarColor") & DblQuote
ConstFile.WriteLine "Const CustomTDLink = " & CBool(rsConfig("CustomTDLink"))	'This is if they choose to use a custom table link color

ConstFile.WriteLine "Const RatingMax = " & rsConfig("RatingMax")
ConstFile.WriteLine "Const PhotosPerRow = " & rsConfig("PhotosPerRow")
ConstFile.WriteLine "Const GroupsPerRow = " & rsConfig("GroupsPerRow")
ConstFile.WriteLine "Const StatTopMax = " & rsConfig("StatTopMax")
ConstFile.WriteLine "Const MemberStatTopMax = " & rsConfig("MemberStatTopMax")
ConstFile.WriteLine "Const CalendarShowBirthdays = " & rsConfig("CalendarShowBirthdays")
ConstFile.WriteLine "Const NewsShowEvents = " & rsConfig("NewsShowEvents")

ConstFile.WriteLine "Const CalendarBirthdayMessage = " & DblQuote & rsConfig("CalendarBirthdayMessage") & DblQuote

ConstFile.WriteLine "Const QuizResult90 = " & DblQuote & rsConfig("QuizResult90") & DblQuote
ConstFile.WriteLine "Const QuizResult60 = " & DblQuote & rsConfig("QuizResult60") & DblQuote
ConstFile.WriteLine "Const QuizResult0 = " & DblQuote & rsConfig("QuizResult0") & DblQuote

ConstFile.WriteLine "Const MediaMegs = " & rsConfig("MediaMegs")
ConstFile.WriteLine "Const PhotosMegs = " & rsConfig("PhotosMegs")

ConstFile.WriteLine "Const SiteMembersOnly = " & rsConfig("SiteMembersOnly")

ConstFile.WriteLine "Const SecureLogin = " & rsConfig("SecureLogin")
ConstFile.WriteLine "Const AllowMemberApplications = " & rsConfig("AllowMemberApplications")
ConstFile.WriteLine "Const MailerFromName = " & DblQuote & rsConfig("MailerFromName") & DblQuote
ConstFile.WriteLine "Const MailerFromAddress = " & DblQuote & rsConfig("MailerFromAddress") & DblQuote
ConstFile.WriteLine "Const MemberNameDisplay = " & DblQuote & rsConfig("MemberNameDisplay") & DblQuote



ConstFile.WriteLine "Const Subdirectory = " & DblQuote & strSubDir & DblQuote
'Put the https address if they choose to use secure logins
if rsConfig("SecureLogin") = 0 then
	if intDomain = 0 then
		ConstFile.WriteLine "Const NonSecurePath = " & DblQuote & "http://www.GroupLoop.com/" & strSubDir & "/" & DblQuote
		ConstFile.WriteLine "Const SecurePath = " & DblQuote & "http://www.GroupLoop.com/" & strSubDir & "/" & DblQuote
	else
		'We are using a domain name, and have a master site with the domain
		if intMasterID > 0 then
			'Take the first subdir out, because then it won't point correctly (ex - bbpc/board will now be board, because www.bbpc.org points to grouploop.com/bbpc)
			SplitDir = Split(strSubDir, "/", -1, 1)
			strSubDir = SplitDir(1)
			ConstFile.WriteLine "Const NonSecurePath = " & DblQuote & strDomainName & "/" & strSubDir & "/" & DblQuote
			ConstFile.WriteLine "Const SecurePath = " & DblQuote & strDomainName & "/" & strSubDir & "/" & DblQuote
		else
			ConstFile.WriteLine "Const NonSecurePath = " & DblQuote & strDomainName & "/" & DblQuote
			ConstFile.WriteLine "Const SecurePath = " & DblQuote & strDomainName & "/" & DblQuote
		end if
	end if
else
	ConstFile.WriteLine "Const NonSecurePath = " & DblQuote & "http://www.OurClubPage.com/" & strSubDir & "/" & DblQuote
	ConstFile.WriteLine "Const SecurePath = " & DblQuote & "https://www.OurClubPage.com/" & strSubDir & "/" & DblQuote
end if

if rsConfig("AllowStore") = 1 then
	ConstFile.WriteLine "Const Returns = " & CBool(rsConfig("Returns"))
	ConstFile.WriteLine "Const TaxRate = " & rsConfig("StoreTax")
	ConstFile.WriteLine "Const IncludeProcessingDays = " & rsConfig("IncludeProcessingDays")
	ConstFile.WriteLine "Const IncludeStoreTax = " & rsConfig("IncludeStoreTax")
	ConstFile.WriteLine "Const SalePriceTitle = " & DblQuote &  rsConfig("SalePriceTitle") & DblQuote
	ConstFile.WriteLine "Const IncludeInventory = " & rsConfig("IncludeInventory")

	ConstFile.WriteLine "Const RateStoreItems = " & rsConfig("RateStoreItems")
	ConstFile.WriteLine "Const RateStoreGroups = " & rsConfig("RateStoreGroups")
	ConstFile.WriteLine "Const ReviewStoreItems = " & rsConfig("ReviewStoreItems")
	ConstFile.WriteLine "Const ReviewStoreGroups = " & rsConfig("ReviewStoreGroups")

	ConstFile.WriteLine "Const CardsAccepted = " & DblQuote &  rsConfig("CardsAccepted") & DblQuote
	ConstFile.WriteLine "Const AutoProcessCards = " & DblQuote &  rsConfig("AutoProcessCards") & DblQuote

end if
ConstFile.WriteLine "'End Constants"

ConstFile.Write "%"
ConstFile.WriteLine ">"



ConstFile.Close
Set ConstFile = Nothing

rsConfig.Close

if not blrsConfigExisted then set rsConfig = Nothing

FileSystem.CopyFile strPath & "constantstemp.inc", strPath & "constants.inc"

if not blFileSystemExisted then set FileSystem = Nothing

%>
