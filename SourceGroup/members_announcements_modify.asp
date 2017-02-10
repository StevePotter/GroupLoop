<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="functions_modify.asp" -->

<%
if not CBool( IncludeAnnouncements ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_announcements_modify.asp")
if not (LoggedAdmin or CBool( AnnouncementsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

blLoggedAdmin = LoggedAdmin()

strTable = "Announcements"
strNoun = "Announcement"
strPluralNoun = "Announcements"
strListSource = "announcements.asp"
strListSourceName = "Announcements"
strModSource = "members_announcements_modify.asp"
strViewSource = "announcements_read.asp?ID="
strViewAction = "Read"
DisplayPrivacy = IncludePrivacy( strTable )

Set rsEdit = Server.CreateObject("ADODB.Recordset")

Sub PrintFields( strRequirements )
	blTemp = ListDupes("announcement", rsEdit)
	PrintItemField False, "Can this " & LCase(strNoun) & " be " & LCase(strViewAction) & " by members only?", "Private", "CheckBox", 0, 0, rsEdit, "Members", DisplayPrivacy, strRequirements
	PrintItemField False, "Date Added", "Date", "DateTime", 0, 0, rsEdit, "Administrators", True, strRequirements
	PrintItemField True, "Subject", "Subject", "Text", 50, 0, rsEdit, "Members", True, strRequirements
	PrintItemField True, "Details", "Body", "TextArea", 50, 30, rsEdit, "Members", True, strRequirements
End Sub

Sub UpdateItemFields
	if Request(rsEdit("ID")) = "1" then

		rsEdit("Private") = GetCheckedResult(Request("Private"))
		rsEdit("Subject") = Format( Request("Subject") )
		if blLoggedAdmin then rsEdit("Date") = AssembleDate( "Date" )
		rsEdit("Body") = GetTextArea( Request("Body") )
		rsEdit("IP") = Request.ServerVariables("REMOTE_HOST")
		rsEdit("ModifiedID") = Session("MemberID")
		rsEdit.Update
	end if
End Sub

GoModify

Set rsEdit = Nothing
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->