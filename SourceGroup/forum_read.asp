<!-- #include file="forum_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeForum ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=ForumTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if


Public DisplayDate, DisplaySubject

Query = "SELECT DisplayDateItemForum, DisplaySubjectItemForum  FROM Look WHERE CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayDate = CBool(rsItem("DisplayDateItemForum"))
	DisplaySubject = CBool(rsItem("DisplaySubjectItemForum"))
rsItem.Close


'Open up the item
Query = "SELECT ID, Date, BaseID, CustomerID, MemberID, CategoryID, Author, Email, Subject, Body, Private FROM ForumMessages WHERE ID = " & intID & " AND CustomerID = " & CustomerID
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The message does not exist.  If you pasted a link, there may be a typo, or the message may have been deleted.  Please refer to the forum to find the desired message, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=forum_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateForum = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Message</a></p>
<%
	AddRating intID, "ForumMessages"
else
	IncrementStat "ForumMessagesRead"
	IncrementHits intID, "ForumMessages"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	intCategoryID = rsItem("CategoryID")
	intBaseID = rsItem("BaseID")

	GetCategory intCategoryID, strName, blPrivate, blMembersOnly
'------------------------End Code-----------------------------
%>
	</p>
	<table width="100%">
		<tr>
			<td align="left">
				<span class="Heading">Topic: <%=strName%> 
<%
'-----------------------Begin Code----------------------------			
				if blMembersOnly then
					%></span><font size="-2">(only members may post messages)</font><%
				else
					%></span><%
				end if
'------------------------End Code-----------------------------
%>
			</td>
<%
			if NeedCategoryMenu("ForumCategories") then
%>
			<td align="right">
				<form action="forum.asp" method="post">
					<font size="-1">Change Topic To:</font><br>
					<% PrintCategoryPullDown intCategoryID %>
					<input type="Submit" value="Switch">
				</form>
			</td>
<%
			end if
%>
		</tr>
	</table>

	<span class="LinkText"><a HREF="forum_post.asp?ReplyID=<%=intID%>&CategoryID=<%=intCategoryID%>">Post Reply</a></span><br>
	<br>

<%
	PrintMessage rsItem

	if intBaseID = 0 then intBaseID = intID

	Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating, Body FROM ForumMessages WHERE ID = " & intBaseID
	Set rsBase = Server.CreateObject("ADODB.Recordset")
	rsBase.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating, Body FROM ForumMessages WHERE BaseID = " & intBaseID & " ORDER BY Date"
	Set rsReplies = Server.CreateObject("ADODB.Recordset")
	rsReplies.CacheSize = 20
	rsReplies.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if not rsReplies.EOF then

		Set RepID = rsReplies("ID")
		Set RepItemDate = rsReplies("Date")
		Set RepEmail = rsReplies("Email")
		Set RepAuthor = rsReplies("Author")
		Set RepSubject = rsReplies("Subject")
		Set RepIsPrivate = rsReplies("Private")
		Set RepMemberID = rsReplies("MemberID")
		Set RepTimesRated = rsReplies("TimesRated")
		Set RepTotalRating = rsReplies("TotalRating")

		'Check the email
		if rsBase("MemberID") > 0 then
			strAuthor = GetNickNameLink( rsBase("MemberID") )
		elseif InStr( rsBase("Email"), "@" ) then
			strAuthor = "<a href='mailto:" & rsBase("Email") & "'>" & rsBase("Author") & "</a>"
		else
			strAuthor = rsBase("Author")
		end if

		if rsBase("Private") = 1 and not blPrivate then
			strPrivate = "Private, "
		else
			strPrivate = ""
		end if

		if rsBase("TimesRated") > 0 and RateForum = 1 then
			strRating = ", Rating: " & GetRating( rsBase("TotalRating"), rsBase("TimesRated") )
		else
			strRating = ""
		end if
		%>
		<% PrintNew(rsBase("Date")) %> <a href="#<%=rsBase("ID")%>"><%=rsBase("Subject")%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(rsBase("Date"), 2)%> <%=strRating%> )</font><br>
		<%

		do until rsReplies.EOF
			if RepMemberID > 0 then
				strAuthor = GetNickNameLink( RepMemberID )
			elseif InStr( RepEmail, "@" ) then
				strAuthor = "<a href='mailto:" & RepEmail & "'>" & RepAuthor & "</a>"
			else
				strAuthor = RepAuthor
			end if

			if RepIsPrivate = 1 and not blPrivate then
				strPrivate = "Private, "
			else
				strPrivate = ""
			end if

			if RepTimesRated > 0 and RateForum = 1 then
				strRating = ", Rating: " & GetRating( RepTotalRating, RepTimesRated )
			else
				strRating = ""
			end if
		%>
			&nbsp;&nbsp;&nbsp;&nbsp; <% PrintNew(RepItemDate) %> <a href="#<%=RepID%>"><%=RepSubject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(RepItemDate, 2)%> <%=strRating%> )</font><br>
		<%
			rsReplies.MoveNext
		loop


		Response.Write "<br><br><br><br><br><br><br><br>"

		'Now print out all the messages

		if rsBase("ID") <> intID then PrintMessage rsBase

		rsReplies.MoveFirst
		do until rsReplies.EOF
			PrintMessage rsReplies
			rsReplies.MoveNext
		loop



	end if

	set rsBase = Nothing
	set rsReplies = Nothing

	set rsItem = Nothing

end if



Sub PrintMessage( rsObject )
	intTempID = rsObject("ID")

	'Check the email
	if rsObject("MemberID") > 0 then
		strAuthor = PrintTDLink(GetNickNameLink( rsObject("MemberID") ))
	elseif InStr( rsObject("Email"), "@" ) then
		strAuthor = "<a href='mailto:" & rsObject("Email") & "'>" & PrintTDLink(rsObject("Author")) & "</a>"
	else
		strAuthor = rsObject("Author")
	end if
%>
	<a name="<%=intTempID%>"></a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsObject("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_forum_modify.asp?Submit=Edit&ID=<%=intTempID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=intTempID%>')">Delete</a></td>
		</tr>
		</table>
<%
	end if

	PrintTableHeader 100
%>
	<tr>
		<td colspan="2" class="<% PrintTDMain %>">
		<table width=100% cellspacing=0 cellpadding=0>
		<tr>
		<td class="<% PrintTDMain %>" align="left">Author: <%=strAuthor%></td>
		<% if DisplayDate then	%>
		<td class="<% PrintTDMainSwitch %>" align="right">Date Written: <%=FormatDateTime(rsObject("Date"), 2)%></td>
			<% end if %>		
		</tr>
		</table>

		</td>
	</tr>
<%	if DisplaySubject then	%>
	<tr>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsObject("Subject")%></td>
	</tr>
	<% end if %>		
	</table>
	<br>
<%
	if rsObject("Private") = 1 and not LoggedMember then
%>
		<p>This is a private message.  If you are a member, <a href="login.asp?Source=forum_read.asp&ID=<%=Request("ID")%>&Submit=Read">click here</a> to log in and view it.</p>
<%
	else
		Response.Write rsObject("Body")
	end if
%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateForum = 1 then
		PrintRatingPulldown intTempID, "", "ForumMessages", "forum_read.asp", "message"
		Response.Write "<br>"
	end if
'------------------------End Code-----------------------------
%>
	<br><br><br><br>
<%
End Sub
'------------------------End Code-----------------------------
%>