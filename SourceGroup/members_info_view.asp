<%
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_info_view.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">View Everyone's Info</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
Query = "SELECT DisplayNickNameListMembers, DisplayNickNameListMembers, DisplayPhotoListMembers, DisplayFullNameListMembers, DisplayBirthdayListMembers, " & _
	"DisplayEMailListMembers, DisplayHomeAddressListMembers, DisplaySecondaryAddressListMembers, DisplayBeeperListMembers, DisplayCellPhoneListMembers, " & _
	"DisplayMembershipLevelListMembers FROM Look WHERE CustomerID = " & CustomerID

Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayNickName = CBool(rsPage("DisplayNickNameListMembers"))
	DisplayPhoto = CBool(rsPage("DisplayPhotoListMembers"))
	DisplayFullName = CBool(rsPage("DisplayFullNameListMembers"))
	DisplayBirthday = CBool(rsPage("DisplayBirthdayListMembers"))
	DisplayEMail = CBool(rsPage("DisplayEMailListMembers"))
	DisplayHomeAddress = CBool(rsPage("DisplayHomeAddressListMembers"))
	DisplaySecondaryAddress = CBool(rsPage("DisplaySecondaryAddressListMembers"))
	DisplayBeeper = CBool(rsPage("DisplayBeeperListMembers"))
	DisplayCellPhone = CBool(rsPage("DisplayCellPhoneListMembers"))
	DisplayMembershipLevel = CBool(rsPage("DisplayMembershipLevelListMembers"))
rsPage.Close




Query = "SELECT * FROM Members WHERE CustomerID = " & CustomerID & " ORDER BY LastName"
rsPage.CacheSize = PageSize
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly


blLoggedMember = True
blLoggedAdmin = LoggedAdmin()
%>
<form METHOD="POST" ACTION="<%=SecurePath%>members_info_view.asp">
<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
<input type="hidden" name="Password" value="<%=Session("Password")%>">
<%
PrintPagesHeader
PrintTableHeader 100
%>
	<tr>
<%
		if DisplayNickName then
%>
		<td class="TDHeader"><%=UsernameLabel%></td>
<%
		end if
		if DisplayFullName then
%>
		<td class="TDHeader">Name</td>
<%
		end if
		if DisplayBirthday then
%>
		<td class="TDHeader">Birthday</td>
<%
		end if
		if DisplayEMail then
%>
		<td class="TDHeader">E-Mail</td>
<%
		end if
		if DisplayHomeAddress then
%>
		<td class="TDHeader">Home Address</td>
<%
		end if
		if DisplaySecondaryAddress then
%>
		<td class="TDHeader">Secondary Address</td>
<%
		end if
		if DisplayBeeper then
%>
		<td class="TDHeader">Beeper</td>
<%
		end if
		if DisplayCellPhone then
%>
		<td class="TDHeader">Cell Phone</td>
<%
		end if
		if RateMembers = 1  and ReviewMembers = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif RateMembers = 0  and ReviewMembers = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif RateMembers = 1  and ReviewMembers = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% end if %>	
<%
		if DisplayMembershipLevel then
%>
		<td class="TDHeader">Membership Level</td>
<%
		end if
%>
	</tr>
<%

for j = 1 to rsPage.PageSize
	if not rsPage.EOF then


'------------------------End Code-----------------------------
%>
	<tr>
<%
		if DisplayNickName then
%>
		<td class="<% PrintTDMain %>"><a href="member.asp?ID=<%=rsPage("ID")%>"><%=PrintTDLink(rsPage("NickName"))%></a></td>
<%'-----------------------Begin Code----------------------------
		end if
		if DisplayFullName then
			if rsPage("FirstName") <> "" and (rsPage("PrivateName") = 1 or blLoggedAdmin) then
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>"><%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%></td>
<%'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>">&nbsp;</td>
<%'-----------------------Begin Code----------------------------
			end if
		end if
		if DisplayBirthday then

			if rsPage("Birthdate") <> "" and (rsPage("PrivateBirthdate") = 1 or blLoggedAdmin) then
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>"><%=FormatDateTime(rsPage("Birthdate"), 2)%></td>
<%'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>">&nbsp;</td>
<%'-----------------------Begin Code----------------------------
			end if
		end if
		if DisplayEMail then
			if rsPage("EMail1") <> "" and (rsPage("PrivateEMail") = 1 or blLoggedAdmin) then
'------------------------End Code-----------------------------
%>			
				<td class="<% PrintTDMain %>"><a href="mailto:<%=rsPage("EMail1")%>"><%=PrintTDLink(rsPage("EMail1"))%></a>
<%'-----------------------Begin Code----------------------------
				if not rsPage("EMail2") = "" then
'------------------------End Code-----------------------------
%>
					<br><a href="mailto:<%=rsPage("EMail2")%>"><%=PrintTDLink(rsPage("EMail2"))%></a>
<%'-----------------------Begin Code----------------------------
				end if
				Response.Write("</td>")
			else
'------------------------End Code-----------------------------
%>				<td class="<% PrintTDMain %>">&nbsp;</a></td>
		
<%'-----------------------Begin Code----------------------------
			end if

		end if
		if DisplayHomeAddress then

			if rsPage("HomeStreet") <> "" and rsPage("HomeCity") <> "" and rsPage("HomeState") <> "" and (rsPage("PrivateHome") = 1 or blLoggedAdmin) then
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>"><%=rsPage("HomeStreet")%><br>
					<%=rsPage("HomeCity")%>,&nbsp;<%=rsPage("HomeState")%>&nbsp;<%=rsPage("HomeZip")%>
					<br><%=rsPage("HomePhone")%>
<%'-----------------------Begin Code----------------------------
					if rsPage("HomeCountry") <> "USA" then Response.Write("<br>" & rsPage("HomeCountry") )
'------------------------End Code-----------------------------
%>
				</td>
<%'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>">&nbsp;</td>
<%'-----------------------Begin Code----------------------------
			end if

		end if
		if DisplaySecondaryAddress then

			if rsPage("SecondaryStreet") <> "" and rsPage("SecondaryCity") <> "" and rsPage("SecondaryState") <> "" and (rsPage("PrivateSecondary") = 1 or blLoggedAdmin) then
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>">
				<%=rsPage("SecondaryDescription")%>:<br>
				<%=rsPage("SecondaryStreet")%><br>
					<%=rsPage("SecondaryCity")%>,&nbsp;<%=rsPage("SecondaryState")%>&nbsp;<%=rsPage("SecondaryZip")%>
					<br><%=rsPage("SecondaryPhone")%>
<%'-----------------------Begin Code----------------------------
				if rsPage("SecondaryPExt") <> "" then Response.Write("&nbsp;&nbsp;&nbsp;&nbsp; Ext. " & rsPage("SecondaryPExt") )
				if rsPage("SecondaryCountry") <> "USA" then Response.Write("<br>" & rsPage("SecondaryCountry"))
'------------------------End Code-----------------------------
%>
				</td>
<%'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<td class="<% PrintTDMain %>">&nbsp;</td>
<%'-----------------------Begin Code----------------------------
			end if

		end if
		if DisplayBeeper then
			strBeeper = ""
			if (rsPage("PrivateBeeper") = 1 or blLoggedAdmin) then strBeeper = rsPage("Beeper")
			if strBeeper = "" then strBeeper = "&nbsp;"
%>
			<td class="<% PrintTDMain %>"><%=strBeeper%></td>
<%
		end if
		if DisplayCellPhone then
			strCellPhone = ""
			if (rsPage("PrivateCellPhone") = 1 or blLoggedAdmin) then strCellPhone = rsPage("CellPhone")
			if strCellPhone = "" then strCellPhone = "&nbsp;"
%>
			<td class="<% PrintTDMain %>"><%=strCellPhone%></td>
<%
		end if


'------------------------End Code-----------------------------
			if RateMembers = 1 and ReviewMembers = 0 then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%> 
				<font size="-2"><a href="member.asp?ID=<%=rsPage("ID")%>">Rate</a></font></td>
<%			elseif RateMembers = 0 and ReviewMembers = 1 then
				if ReviewsExist( "Members", rsPage("ID") ) then
%>					<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="member.asp?ID=<%=rsPage("ID")%>">Read/Add Review</a></font></td>
<%				else
%>					<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="member.asp?ID=<%=rsPage("ID")%>">Add Review</a></font></td>
<%				end if
			elseif RateMembers = 1 and ReviewMembers = 1 then
				if ReviewsExist( "Members", rsPage("ID") ) then
%>					<td class="<% PrintTDMain %>" align=center><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%> 
					<font size="-2"><a href="member.asp?ID=<%=rsPage("ID")%>">Rate and Read/Add Review</a></font></td>
<%				else
%>					<td class="<% PrintTDMain %>" align=center><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%> 
					<font size="-2"><a href="member.asp?ID=<%=rsPage("ID")%>">Rate/Add Review</a></font></td>
<%				end if
			end if
		if DisplayMembershipLevel then
			strLevel = "Regular Member"
			if rsPage("Admin") = 1 then strLevel = "Site Administrator"
%>
			<td class="<% PrintTDMain %>"><%=strLevel%></td>	
<%
		end if
%>
	</tr>
<%
'-----------------------Begin Code----------------------------
	ChangeTDMain

	rsPage.MoveNext
end if
next
rsPage.Close
set rsPage = Nothing
Response.Write "</table>"


'Give them the link to change the section's properties
if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
	Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Membership'>Change This Section's Options</a></p>"
end if


'------------------------End Code-----------------------------
%>

