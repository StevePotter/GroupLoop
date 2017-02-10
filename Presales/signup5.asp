<!-- #include file="header.asp" -->
<!-- #include file="dsn.asp" -->
<!-- #include file="functions.asp" -->
<% AddHit "signup5.asp" %>

<p class=Heading align=center>
Step 5. Add Starting Members

<p>You may add up to 4 members to start with.  You don't have to add any if you don't want to, and 
none of this is permanent.  You can always add and remove members after your site is created.  
<b>DO NOT ENTER YOURSELF!!!</b></p>

<%
'We are creating a new child site.. secret!
if Request("ParentID") <> "" then intParentID = Request("ParentID")


if Request("Version") = "" or ( Request("Version") <> "Gold" and Request("Version") <> "Free" and Request("Version") <> "Parent"  ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

strType = Request("Version")


if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the initial look.  Please go through the sign-up process from the beginning."))
intSchemeID = CInt(Request("ID"))

%>


* indicates required information

<script language="JavaScript">
<!--
	function submit_page(form) {
		//Error message variable
		var strError = "";


		for (var i=1; i<=4; i++) {
			//They entered something
			if( form.elements['FirstName'+i].value != "" || form.elements['LastName'+i].value != "" || form.elements['NickName'+i].value != "" || form.elements['Password'+i].value != "" || form.elements['EMail'+i].value != "" ){
				if(form.elements['FirstName'+i].value == "")
					strError += "          You forgot the first name for #" + i + ". \n";
				else{
					if (!ValidateStuff(form.elements['FirstName'+i].value))
						strError += "          You may not enter any of these characters for the first name for " + i + " - \\/:*?\"<>|. \n";
				}
				if(form.elements['LastName'+i].value == "")
					strError += "          You forgot the last Name for #" + i + ". \n";
				else{
					if (!ValidateStuff(form.elements['LastName'+i].value))
						strError += "          You may not enter any of these characters for the last name for " + i + " - \\/:*?\"<>|. \n";
				}
				if(form.elements['NickName'+i].value == "")
					strError += "          You forgot the nickname for #" + i + ". \n";
				else{
					if (!ValidateStuff(form.elements['NickName'+i].value))
						strError += "          You may not enter any of these characters for the nickname for " + i + " - \\/:*?\"<>|. \n";
				}
				if(form.elements['Password'+i].value == "")
					strError += "          You forgot the password for #" + i + ". \n";
				else{
					if (!ValidateStuff(form.elements['Password'+i].value))
						strError += "          You may not enter any of these characters for the password for " + i + " - \\/:*?\"<>|. \n";
				}
				if(form.elements['EMail'+i].value == "")
					strError += "          You forgot the e-mail for #" + i + ". \n";
				else{
					if ((getFront(form.elements['EMail'+i].value,"@") == null) || (getEnd(form.elements['EMail'+i].value,"@") == ""))
						strError += "          Please enter a valid e-mail address, such as JoesPizza@aol.com for #" + i + ". \n";
					else{
						if (!ValidateStuff(form.elements['EMail'+i].value))
							strError += "          You may not enter any of these characters for the EMail for " + i + " - \\/:*?\"<>|. \n";
					}
				}
			}
		}
		if(strError == "") {
			return true;
		}
        else{
			strError = "Sorry, but you must go back and fix the following errors before you can proceed: \n" + strError;
			alert (strError);
			return false;
		}   
	}

	function getFront(mainStr,searchStr){
		foundOffset = mainStr.indexOf(searchStr)
        if (foundOffset <= 0) {
            return null // if the @ symbol is missing the value is -1
                        // if the @ symbol is the first char the value is 0
        } 
        else {
            return mainStr.substring(0,foundOffset)
        }
    }
    
    function getEnd(mainStr,searchStr) {
        foundOffset = mainStr.indexOf(searchStr)
        if (foundOffset <= 0) {
            return ""   // if the @ symbol is missing the value is -1
                        // if the @ symbol is the first char the value is 0
        }
        else {
            return mainStr.substring(foundOffset+searchStr.length,mainStr.length)
        }
    }

	function ValidateStuff(string) {
		var Invalid='\\/:/*?\"<>|'

		for (var i=0; i<string.length; i++) {
			if (Invalid.indexOf(string.charAt(i)) >= 0) {
				return false;
			}
		}

		return true;
	} 
//-->
</SCRIPT>

<form METHOD="POST" ACTION="signup6.asp" name="Signup" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
<%
	if intParentID <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
		<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
	end if
%>
	<input type="hidden" name="SessionID" value="<%=Session("SessionID")%>">
	<input type="hidden" name="SiteMembersOnly" value="<%=Request("SiteMembersOnly")%>">
	<input type="hidden" name="AllowMemberApplications" value="<%=Request("AllowMemberApplications")%>">
	<input type="hidden" name="IncludeNewsletter" value="<%=Request("IncludeNewsletter")%>">
	<input type="hidden" name="NewsletterMembers" value="<%=Request("NewsletterMembers")%>">
	<input type="hidden" name="IncludeAnnouncements" value="<%=Request("IncludeAnnouncements")%>">
	<input type="hidden" name="RateAnnouncements" value="<%=Request("RateAnnouncements")%>">
	<input type="hidden" name="ReviewAnnouncements" value="<%=Request("ReviewAnnouncements")%>">
	<input type="hidden" name="IncludeMeetings" value="<%=Request("IncludeMeetings")%>">
	<input type="hidden" name="MeetingsMembers" value="<%=Request("MeetingsMembers")%>">
	<input type="hidden" name="RateMeetings" value="<%=Request("RateMeetings")%>">
	<input type="hidden" name="ReviewMeetings" value="<%=Request("ReviewMeetings")%>">
	<input type="hidden" name="IncludeStories" value="<%=Request("IncludeStories")%>">
	<input type="hidden" name="RateStories" value="<%=Request("RateStories")%>">
	<input type="hidden" name="ReviewStories" value="<%=Request("ReviewStories")%>">
	<input type="hidden" name="IncludeCalendar" value="<%=Request("IncludeCalendar")%>">
	<input type="hidden" name="CalendarShowBirthdays" value="<%=Request("CalendarShowBirthdays")%>">
	<input type="hidden" name="RateCalendar" value="<%=Request("RateCalendar")%>">
	<input type="hidden" name="ReviewCalendar" value="<%=Request("ReviewCalendar")%>">
	<input type="hidden" name="IncludeLinks" value="<%=Request("IncludeLinks")%>">
	<input type="hidden" name="RateLinks" value="<%=Request("RateLinks")%>">
	<input type="hidden" name="ReviewLinks" value="<%=Request("ReviewLinks")%>">
	<input type="hidden" name="IncludeQuotes" value="<%=Request("IncludeQuotes")%>">
	<input type="hidden" name="RateQuotes" value="<%=Request("RateQuotes")%>">
	<input type="hidden" name="ReviewQuotes" value="<%=Request("ReviewQuotes")%>">
	<input type="hidden" name="IncludeQuizzes" value="<%=Request("IncludeQuizzes")%>">
	<input type="hidden" name="QuizzesMembers" value="<%=Request("QuizzesMembers")%>">
	<input type="hidden" name="RateQuizzes" value="<%=Request("RateQuizzes")%>">
	<input type="hidden" name="ReviewQuizzes" value="<%=Request("ReviewQuizzes")%>">
	<input type="hidden" name="IncludeVoting" value="<%=Request("IncludeVoting")%>">
	<input type="hidden" name="VotingMembers" value="<%=Request("VotingMembers")%>">
	<input type="hidden" name="RateVoting" value="<%=Request("RateVoting")%>">
	<input type="hidden" name="ReviewVoting" value="<%=Request("ReviewVoting")%>">
	<input type="hidden" name="IncludePhotos" value="<%=Request("IncludePhotos")%>">
	<input type="hidden" name="PhotosMembers" value="<%=Request("PhotosMembers")%>">
	<input type="hidden" name="RatePhotos" value="<%=Request("RatePhotos")%>">
	<input type="hidden" name="IncludePhotoCaptions" value="<%=Request("IncludePhotoCaptions")%>">
	<input type="hidden" name="IncludeForum" value="<%=Request("IncludeForum")%>">
	<input type="hidden" name="RateForum" value="<%=Request("RateForum")%>">
	<input type="hidden" name="IncludeGuestbook" value="<%=Request("IncludeGuestbook")%>">
	<input type="hidden" name="RateGuestbook" value="<%=Request("RateGuestbook")%>">
	<input type="hidden" name="ReviewGuestbook" value="<%=Request("ReviewGuestbook")%>">
	<input type="hidden" name="IncludeMedia" value="<%=Request("IncludeMedia")%>">
	<input type="hidden" name="MediaMembers" value="<%=Request("MediaMembers")%>">
	<input type="hidden" name="RateMedia" value="<%=Request("RateMedia")%>">
	<input type="hidden" name="ReviewMedia" value="<%=Request("ReviewMedia")%>">

	<input type="hidden" name="ID" value="<%=intSchemeID%>">
	<input type="hidden" name="Version" value="<%=strType%>">

	<% PrintTableHeader 0 %>

	<tr>
		<td class=TDHeader align=center>
			&nbsp;	
		</td>
		<td class=TDHeader align=center>
			* First Name
		</td>
		<td class=TDHeader align=center>
			* Last Name
		</td>
		<td class=TDHeader align=center>
			* NickName
		</td>
		<td class=TDHeader align=center>
			* Initial Password
		</td>
		<td class=TDHeader align=center>
			* E-Mail Address
		</td>
	</tr>
<%
for i = 1 to 4
%>
	<tr>
		<td class="<% PrintTDMain %>" align="left">
			<%=i%>.
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="FirstName<%=i%>" size="10" maxlength="100">
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="LastName<%=i%>" size="10" maxlength="100">
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="NickName<%=i%>" size="10" maxlength="100">
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Password<%=i%>" size="10" maxlength="100">
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<input type="text" name="EMail<%=i%>" size="20" maxlength="200">
		</td>
	</tr>
<%
next
%>
	<tr>
		<td class=<% PrintTDMain %> align=center colspan=6>
			<input type="submit" name="Submit" value="I'm Done" >
		</td>
	</tr>
</table>

</form>

<!-- #include file="closedsn.asp" -->

<!-- #include file="footer.asp" -->