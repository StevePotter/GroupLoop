<!-- #include file="header.asp" -->
<!-- #include file="dsn.asp" -->
<!-- #include file="functions.asp" -->
<% AddHit "signup6.asp" %>

<p class=Heading align=center>
Step 6. Enter Your Site 
<%
'We are creating a new child site.. secret!
if Request("ParentID") <> "" then intParentID = Request("ParentID")

if Request("Version") = "Gold" or Request("Version") = "Parent" then Response.Write " And Billing"
%>
 Info</p>

<%
if Request("Version") = "" or ( Request("Version") <> "Gold" and Request("Version") <> "Free" and Request("Version") <> "Parent"  ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

strType = Request("Version")


if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the initial look.  Please go through the sign-up process from the beginning."))
intSchemeID = CInt(Request("ID"))

%>
<p>The information you provide is strictly confidential and used for our records only.</p>

* indicates required information

<script language="JavaScript">
<!--
	//Throw out all the stuff we don't want ($)
	function ConvertInt(currCheck) {
		if (!currCheck) return '';
		for (var i=0, currOutput='', valid="0123456789"; i<currCheck.length; i++)
			if (valid.indexOf(currCheck.charAt(i)) != -1)
				currOutput += currCheck.charAt(i);
		return currOutput;
	}


	function submit_page(form) {
		//Error message variable
		var strError = "";

		form.SalesmanID.value = ConvertInt(form.SalesmanID.value);

        if(form.FirstName.value == "")
			strError += "          You forgot your First Name. \n";
        if(form.LastName.value == "")
			strError += "          You forgot your Last Name. \n";
        if(form.NickName.value == "")
			strError += "          You forgot your NickName. \n";
        if(form.EMail.value == "")
			strError += "          You forgot your EMail. \n";
		else{
			if ((getFront(form.EMail.value,"@") == null) || (getEnd(form.EMail.value,"@") == ""))
				strError += "          Please enter a valid e-mail address, such as JoesPizza@aol.com. \n";
		}
        if(form.PW1.value == "" || form.PW2.value == "")
			strError += "          You forgot your Password. \n";
        if(form.PW1.value != form.PW2.value)
			strError += "          The passwords you typed were not exactly your same.  Please retype yourn. \n";
        if(form.Title.value == "")
			strError += "          You forgot your Site Title. \n";
<%
	if strType = "Gold" or strType = "Parent" then
%>
        if(form.Street1.value == "")
			strError += "          You forgot your Street Address. \n";
        if(form.City.value == "")
			strError += "          You forgot your City. \n";
        if(form.Zip.value == "")
			strError += "          You forgot your Zip Code. \n";
        if(form.State.value == "" && (form.Country.value == "USA" || form.Country.value == "CAN"))
			strError += "          You forgot your State. \n";
        if(form.Phone.value == "")
			strError += "          You forgot your Phone Number. \n";

		//Need to choose the domain name
		if (!form.UseDomain[0].checked && !form.UseDomain[1].checked)
			strError += "          You forgot to choose whether or not you wish to use a domain name. \n";
		//Need to enter the domain name
		if (form.UseDomain[0].checked && form.DomainName.value == "")
			strError += "          You forgot your Domain Name. \n";
		//Need to enter the domain name
		if (form.UseDomain[0].checked && form.DomainAction.value == "")
			strError += "          You forgot to choose if your domain is a New or Transfer domain. \n";
		//Need to enter the domain name
		if (form.UseDomain[1].checked && form.SubDirectory.value == "")
			strError += "          You forgot your Sub-Directory. \n";
		if (form.UseDomain[1].checked && !ValidateDir(form.SubDirectory.value))
			strError += "          You entered a bad sub-directory name.  The name cannot include these characters: .\\/:*?\"<>|\n";

		//They didn't enter a name or company
        if(form.CCFirstName.value == "" && form.CCLastName.value == "" && form.CCCompany.value == "")
			strError += "          You forgot your Credit Card Name or Company. \n";
		//They entered a first name, but not a last
		else if( form.CCFirstName.value != "" && form.CCLastName.value == "" )
			strError += "          You forgot your Credit Card Last Name. \n";
		//They entered a last name, but not a first
		else if( form.CCFirstName.value == "" && form.CCLastName.value != "" )
			strError += "          You forgot your Credit Card First Name. \n";

        if(form.CCNumber.value == "")
			strError += "          You forgot your Credit Card Number. \n";
<%
	else
%>
		if (form.SubDirectory.value == "")
			strError += "          You forgot your Sub-Directory. \n";
		else
			if (!ValidateDir(form.SubDirectory.value))
				strError += "          You entered a bad sub-directory name.  The name cannot include spaces or these characters: \\  /  :  *  ?  \"  <  >  |  .  \'\n";
<%
	end if
%>
        if(!form.Agree.checked)
			strError += "          You forgot to check the Agree box at the bottom. \n";

		if(strError == "") {
			alert('Get ready, because your site will be ready momentarily.  Thank you, and enjoy your new site!');
			return true;
		}
        else{
			strError = "Sorry, but you must go back and fix the following errors before you can sign up: \n" + strError;
			alert (strError);
			return false;
		}   
	}

	

	function ValidateDir(string) {
		var Invalid='\\/:/*?\"\'<>|. '

		for (var i=0; i<string.length; i++) {
			if (Invalid.indexOf(string.charAt(i)) >= 0) {
				return false;
			}
		}

		return true;
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
//-->
</SCRIPT>

<form METHOD="POST" ACTION="process.asp" name="Signup" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
<%
	if intParentID <> "" then
%>
		<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
		<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
	end if
%>
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
<%
	for i = 1 to 4
%>
	<input type="hidden" name="FirstName<%=i%>" value="<%=Request("FirstName" & i)%>">
	<input type="hidden" name="LastName<%=i%>" value="<%=Request("LastName" & i)%>">
	<input type="hidden" name="NickName<%=i%>" value="<%=Request("NickName" & i)%>">
	<input type="hidden" name="Password<%=i%>" value="<%=Request("Password" & i)%>">
	<input type="hidden" name="EMail<%=i%>" value="<%=Request("EMail" & i)%>">
<%
	next
%>
	<input type="hidden" name="SchemeID" value="<%=intSchemeID%>">
	<input type="hidden" name="Version" value="<%=strType%>">
	<table width="100%" cellspacing=2 cellpadding=1 border=0>

	<tr>
		<td class=TDHeader align=center colspan=2>
			Your Information
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* First Name
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="FirstName" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Last Name
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="LastName" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Your NickName For Logging Into Your Site
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="NickName" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* E-Mail Address
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="EMail" size="40">
		</td>
	</tr>

	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Password For Site
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="password" name="PW1" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Confirm Password
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="password" name="PW2" size="40">
		</td>
	</tr>

<%
if strType = "Gold" or strType = "Parent" then
	strReq = "*"
else
	strReq = ""
end if
%>

	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strReq%> Street Address
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Street1" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Street Address Line 2
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Street2" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strReq%> City
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="City" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			State/Province
		</td>
		<td class="<% PrintTDMain %>">
			<% PrintStatesProvinces "State" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strReq%> Zip Code
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Zip" size="8">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strReq%> Country
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<%PrintCountry "Country"%>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strReq%> Phone Number (xxx.xxx.xxxx)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Phone" size="12">
		</td>
	</tr>



	<tr>
		<td class=TDHeader align=center colspan=2>
			Site Information
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Site Title (ex - 'Rutgers Fishing Club')
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Title" size="50">
		</td>
	</tr>

<%
	if strType = "Gold" or strType = "Parent" then
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Organization you are part of (optional, ex - 'Virginia Tech')
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Organization" size="50" onChange="this.form.DomainAction.selectedIndex = 0;">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Are you going to use a domain name?  (www.yoursite.com instead of www.GroupLoop.com/yoursite)<br>
			There is a $50 <b>non-refundable</b> setup fee and a $2/month service fee in addition to any fees from the registrar.
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="radio" name="UseDomain" value="1" onClick="this.form.SubDirectory.disabled=true; this.form.SubDirectory.value=''; this.form.DomainName.disabled=false; this.form.DomainAction.disabled=false; this.form.DomainName.focus();" >
			Yes 
			<input type="radio" name="UseDomain" value="0" onClick="this.form.DomainAction.selectedIndex = 0; this.form.SubDirectory.disabled=false; this.form.DomainName.value=''; this.form.DomainAction.disabled=true; this.form.DomainName.disabled=true; this.form.SubDirectory.focus();" >
			No 
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			If No, what would you like your Sub-Directory to be?  For example, if you want your address 
			to be www.GroupLoop.com/TheSmiths, enter 'TheSmiths'.
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="SubDirectory" size="40" onKeyUp="if (this.disabled) this.value='';" >
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			If Yes, what is the domain name you will use (please include the www.  ex - 'www.joesfriends.net')
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="DomainName" size="50" onKeyUp="if (this.disabled) this.value='';" onFocus="window.temp=this.value" onBlur="if (window.temp != this.value) this.form.DomainAction.focus();" onChange="this.form.DomainAction.focus();">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			If Yes, is this a New domain or does it need to be Transferred (it's a Transfer if you've already 
			bought it).  If it is New, we will sign up for you with Register.com.  You should receive 
			a separate charge from them.  Please note that registering and transferring domain names 
			can take a while (not our fault, we hate it too).
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<select name="DomainAction" size=1>
				<option value=""></option>
				<option value="New">New</option>
				<option value="Transfer">Transfer</option>
			</select>
		</td>
	</tr>
<%
	else
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Organization you are part of (optional, ex - 'Virginia Tech')
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Organization" size="50">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			What would you like your Sub-Directory to be?  For example, if you want your address 
			to be www.GroupLoop.com/TheSmiths, enter 'TheSmiths'.
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="SubDirectory" size="40">
		</td>
	</tr>
<%
	end if
%>
	<tr>
		<td class=TDHeader align=center colspan=2>
			Salesman Information
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align=center colspan=2>
			If you heard about us through a salesman, please enter the salesman number they gave you below.  Do not 
			enter their name.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Salesman ID Number
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="SalesmanID" size="4">
		</td>
	</tr>

<%
	if strType = "Gold" or strType = "Parent" then
%>


	<tr>
		<td class=TDHeader align=center colspan=2>
			Billing Information (information must be exact for verification)
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Name On Card (First then Last)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCFirstName" size="20">&nbsp;
			<input type="text" name="CCLastName" size="20">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Company Name (optional)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCCompany" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Card Type
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<select name="CCType" size=1>
				<option value="VISA">VISA</option>
				<option value="MasterCard">MasterCard</option>
				<option value="AmEx">American Express</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Card Number (no dashes or spaces please)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCNumber" size="18">
		</td>
	</tr>
    <tr> 
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Expiration Date
		</td>
		<td class="<% PrintTDMain %>" align="left">
        <select name="CCExpMonth">
			<option value="1">January</option>
			<option value="2">February</option>
			<option value="3">March</option>
			<option value="4">April</option>
			<option value="5">May</option>
			<option value="6">June</option>
			<option value="7">July</option>
			<option value="8">August</option>
			<option value="9">September</option>
			<option value="10">October</option>
			<option value="11">November</option>
			<option value="12">December</option>
        </select>
        <select name="CCExpYear">
			<option value="2001">2001</option>
			<option value="2002">2002</option>
			<option value="2003">2003</option>
			<option value="2004">2004</option>
			<option value="2005">2005</option>
			<option value="2006">2006</option>
			<option value="2007">2007</option>
			<option value="2008">2008</option>
			<option value="2009">2009</option>
			<option value="2010">2010</option>
        </select>
		</td>
    </tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Street Address For Card (leave the address stuff blank if it is the same address 
			you entered above)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCStreet1" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Street Address Line 2
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCStreet2" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			City
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCCity" size="40">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			State/Province
		</td>
		<td class="<% PrintTDMain %>">
			<% PrintStatesProvinces "CCState" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Zip Code
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="CCZip" size="8">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Country
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<%PrintCountry "CCCountry"%>
		</td>
	</tr>
<%
	end if
%>
	<tr>
		<td class=<% PrintTDMain %> align=center colspan=2>
			I have read and agree to the Terms Of Service <input type="checkbox" name="Agree" value="Yes">
		</td>
	</tr>
	<tr>
		<td class=<% PrintTDMain %> align=center colspan=2>
			<input type="submit" name="Submit" value="Create My Site" >
		</td>
	</tr>
</table>
</form>

<%
Sub PrintCountry( strName )
%>
	<SELECT Name="<%=strName%>" size="1">
	<OPTION VALUE="AFG"> Afghanistan
	<OPTION VALUE="ALB"> Albania
	<OPTION VALUE="DZA"> Algeria
	<OPTION VALUE="ASM"> American Samoa
	<OPTION VALUE="AND"> Andorra
	<OPTION VALUE="AGO"> Angola
	<OPTION VALUE="AIA"> Anguilla
	<OPTION VALUE="ATA"> Antarctica
	<OPTION VALUE="ATG"> Antigua and Barbuda
	<OPTION VALUE="ARG"> Argentina
	<OPTION VALUE="ARM"> Armenia
	<OPTION VALUE="ABW"> Aruba
	<OPTION VALUE="AUS"> Australia
	<OPTION VALUE="AUT"> Austria
	<OPTION VALUE="AZE"> Azerbaijan
	<OPTION VALUE="BHS"> Bahamas
	<OPTION VALUE="BHR"> Bahrain
	<OPTION VALUE="BGD"> Bangladesh
	<OPTION VALUE="BRB"> Barbados
	<OPTION VALUE="BLR"> Belarus
	<OPTION VALUE="BEL"> Belgium
	<OPTION VALUE="BLZ"> Belize
	<OPTION VALUE="BEN"> Benin
	<OPTION VALUE="BMU"> Bermuda
	<OPTION VALUE="BTN"> Bhutan
	<OPTION VALUE="BOL"> Bolivia
	<OPTION VALUE="BIH"> Bosnia And Herzegowina
	<OPTION VALUE="BWA"> Botswana
	<OPTION VALUE="BVT"> Bouvet Island
	<OPTION VALUE="BRA"> Brazil
	<OPTION VALUE="IOT"> British Indian Ocean Territory
	<OPTION VALUE="BRN"> Brunei Darussalam
	<OPTION VALUE="BGR"> Bulgaria
	<OPTION VALUE="BFA"> Burkina Faso
	<OPTION VALUE="BDI"> Burundi
	<OPTION VALUE="KHM"> Cambodia
	<OPTION VALUE="CMR"> Cameroon
	<OPTION VALUE="CAN"> Canada
	<OPTION VALUE="CPV"> Cape Verde
	<OPTION VALUE="CYM"> Cayman Islands
	<OPTION VALUE="CAF"> Central African Republic
	<OPTION VALUE="TCD"> Chad
	<OPTION VALUE="CHL"> Chile
	<OPTION VALUE="CHN"> China
	<OPTION VALUE="CXR"> Christmas Island
	<OPTION VALUE="CCK"> Cocos (Keeling) Islands
	<OPTION VALUE="COL"> Colombia
	<OPTION VALUE="COM"> Comoros
	<OPTION VALUE="COG"> Congo
	<OPTION VALUE="COK"> Cook Islands
	<OPTION VALUE="CRI"> Costa Rica
	<OPTION VALUE="CIV"> Cote D Ivoire
	<OPTION VALUE="HRV"> Croatia (Hrvatska)
	<OPTION VALUE="CYP"> Cyprus
	<OPTION VALUE="CZE"> Czech Republic
	<OPTION VALUE="DNK"> Denmark
	<OPTION VALUE="DJI"> Djibouti
	<OPTION VALUE="DMA"> Dominica
	<OPTION VALUE="DOM"> Dominican Republic
	<OPTION VALUE="TMP"> East Timor
	<OPTION VALUE="ECU"> Ecuador
	<OPTION VALUE="EGY"> Egypt
	<OPTION VALUE="SLV"> El Salvador
	<OPTION VALUE="GNQ"> Equatorial Guinea
	<OPTION VALUE="ERI"> Eritrea
	<OPTION VALUE="EST"> Estonia
	<OPTION VALUE="ETH"> Ethiopia
	<OPTION VALUE="FLK"> Falkland Islands (Malvinas)
	<OPTION VALUE="FRO"> Faroe Islands
	<OPTION VALUE="FJI"> Fiji
	<OPTION VALUE="FIN"> Finland
	<OPTION VALUE="FRA"> France
	<OPTION VALUE="FXX"> France, Metropolitan
	<OPTION VALUE="GUF"> French Guiana
	<OPTION VALUE="PYF"> French Polynesia
	<OPTION VALUE="ATF"> French Southern Territories
	<OPTION VALUE="GAB"> Gabon
	<OPTION VALUE="GMB"> Gambia
	<OPTION VALUE="GEO"> Georgia
	<OPTION VALUE="DEU"> Germany
	<OPTION VALUE="GHA"> Ghana
	<OPTION VALUE="GIB"> Gibraltar
	<OPTION VALUE="GRC"> Greece
	<OPTION VALUE="GRL"> Greenland
	<OPTION VALUE="GRD"> Grenada
	<OPTION VALUE="GLP"> Guadeloupe
	<OPTION VALUE="GUM"> Guam
	<OPTION VALUE="GTM"> Guatemala
	<OPTION VALUE="GIN"> Guinea
	<OPTION VALUE="GNB"> Guinea-Bissau
	<OPTION VALUE="GUY"> Guyana
	<OPTION VALUE="HTI"> Haiti
	<OPTION VALUE="HMD"> Heard And McDonald Islands
	<OPTION VALUE="HND"> Honduras
	<OPTION VALUE="HKG"> Hong Kong
	<OPTION VALUE="HUN"> Hungary
	<OPTION VALUE="ISL"> Iceland
	<OPTION VALUE="IND"> India
	<OPTION VALUE="IDN"> Indonesia
	<OPTION VALUE="IRL"> Ireland
	<OPTION VALUE="ISR"> Israel
	<OPTION VALUE="ITA"> Italy
	<OPTION VALUE="JAM"> Jamaica
	<OPTION VALUE="JPN"> Japan
	<OPTION VALUE="JOR"> Jordan
	<OPTION VALUE="KAZ"> Kazakhstan
	<OPTION VALUE="KEN"> Kenya
	<OPTION VALUE="KIR"> Kiribati
	<OPTION VALUE="PRK"> Korea, Democratic People's Republic Of
	<OPTION VALUE="KOR"> Korea, Republic Of
	<OPTION VALUE="KWT"> Kuwait
	<OPTION VALUE="KGZ"> Kyrgyzstan
	<OPTION VALUE="LAO"> Lao People's Democratic Republic
	<OPTION VALUE="LVA"> Latvia
	<OPTION VALUE="LBN"> Lebanon
	<OPTION VALUE="LSO"> Lesotho
	<OPTION VALUE="LBR"> Liberia
	<OPTION VALUE="LIE"> Liechtenstein
	<OPTION VALUE="LTU"> Lithuania
	<OPTION VALUE="LUX"> Luxembourg
	<OPTION VALUE="MAC"> Macau
	<OPTION VALUE="MKD"> Macedonia, Former Yugoslav Republic Of
	<OPTION VALUE="MDG"> Madagascar
	<OPTION VALUE="MWI"> Malawi
	<OPTION VALUE="MYS"> Malaysia
	<OPTION VALUE="MDV"> Maldives
	<OPTION VALUE="MLI"> Mali
	<OPTION VALUE="MLT"> Malta
	<OPTION VALUE="MHL"> Marshall Islands
	<OPTION VALUE="MTQ"> Martinique
	<OPTION VALUE="MRT"> Mauritania
	<OPTION VALUE="MUS"> Mauritius
	<OPTION VALUE="MYT"> Mayotte
	<OPTION VALUE="MEX"> Mexico
	<OPTION VALUE="FSM"> Micronesia, Federated States Of
	<OPTION VALUE="MDA"> Moldova, Republic Of
	<OPTION VALUE="MCO"> Monaco
	<OPTION VALUE="MNG"> Mongolia
	<OPTION VALUE="MSR"> Montserrat
	<OPTION VALUE="MAR"> Morocco
	<OPTION VALUE="MOZ"> Mozambique
	<OPTION VALUE="MMR"> Myanmar
	<OPTION VALUE="NAM"> Namibia
	<OPTION VALUE="NRU"> Nauru
	<OPTION VALUE="NPL"> Nepal
	<OPTION VALUE="NLD"> Netherlands
	<OPTION VALUE="ANT"> Netherlands Antilles
	<OPTION VALUE="NCL"> New Caledonia
	<OPTION VALUE="NZL"> New Zealand
	<OPTION VALUE="NIC"> Nicaragua
	<OPTION VALUE="NER"> Niger
	<OPTION VALUE="NGA"> Nigeria
	<OPTION VALUE="NIU"> Niue
	<OPTION VALUE="NFK"> Norfolk Island
	<OPTION VALUE="MNP"> Northern Mariana Islands
	<OPTION VALUE="NOR"> Norway
	<OPTION VALUE="OMN"> Oman
	<OPTION VALUE="PAK"> Pakistan
	<OPTION VALUE="PLW"> Palau
	<OPTION VALUE="PAN"> Panama
	<OPTION VALUE="PNG"> Papua New Guinea
	<OPTION VALUE="PRY"> Paraguay
	<OPTION VALUE="PER"> Peru
	<OPTION VALUE="PHL"> Philippines
	<OPTION VALUE="PCN"> Pitcairn
	<OPTION VALUE="POL"> Poland
	<OPTION VALUE="PRT"> Portugal
	<OPTION VALUE="PRI"> Puerto Rico
	<OPTION VALUE="QAT"> Qatar
	<OPTION VALUE="REU"> Reunion
	<OPTION VALUE="ROM"> Romania
	<OPTION VALUE="RUS"> Russian Federation
	<OPTION VALUE="RWA"> Rwanda
	<OPTION VALUE="KNA"> Saint Kitts And Nevis
	<OPTION VALUE="LCA"> Saint Lucia
	<OPTION VALUE="VCT"> Saint Vincent  Grenadines
	<OPTION VALUE="WSM"> Samoa
	<OPTION VALUE="SMR"> San Marino
	<OPTION VALUE="STP"> Sao Tome And Principe
	<OPTION VALUE="SAU"> Saudi Arabia
	<OPTION VALUE="SEN"> Senegal
	<OPTION VALUE="SYC"> Seychelles
	<OPTION VALUE="SLE"> Sierra Leone
	<OPTION VALUE="SGP"> Singapore
	<OPTION VALUE="SVK"> Slovakia (Slovak Republic)
	<OPTION VALUE="SVN"> Slovenia
	<OPTION VALUE="SLB"> Solomon Islands
	<OPTION VALUE="SOM"> Somalia
	<OPTION VALUE="ZAF"> South Africa
	<OPTION VALUE="SGS"> South Georgia  Sandwich Islands
	<OPTION VALUE="ESP"> Spain
	<OPTION VALUE="LKA"> Sri Lanka
	<OPTION VALUE="SHN"> St. Helena
	<OPTION VALUE="SPM"> St. Pierre And Miquelon
	<OPTION VALUE="SUR"> Suriname
	<OPTION VALUE="SJM"> Svalbard And Jan Mayen Islands
	<OPTION VALUE="SWZ"> Swaziland
	<OPTION VALUE="SWE"> Sweden
	<OPTION VALUE="CHE"> Switzerland
	<OPTION VALUE="TWN"> Taiwan
	<OPTION VALUE="TJK"> Tajikistan
	<OPTION VALUE="TZA"> Tanzania, United Republic Of
	<OPTION VALUE="THA"> Thailand
	<OPTION VALUE="TGO"> Togo
	<OPTION VALUE="TKL"> Tokelau
	<OPTION VALUE="TON"> Tonga
	<OPTION VALUE="TTO"> Trinidad And Tobago
	<OPTION VALUE="TUN"> Tunisia
	<OPTION VALUE="TUR"> Turkey
	<OPTION VALUE="TKM"> Turkmenistan
	<OPTION VALUE="TCA"> Turks And Caicos Islands
	<OPTION VALUE="TUV"> Tuvalu
	<OPTION VALUE="UGA"> Uganda
	<OPTION VALUE="UKR"> Ukraine
	<OPTION VALUE="ARE"> United Arab Emirates
	<OPTION VALUE="GBR"> United Kingdom
	<OPTION VALUE="USA" SELECTED> United States
	<OPTION VALUE="UMI"> United States Minor Outlying Islands
	<OPTION VALUE="URY"> Uruguay
	<OPTION VALUE="UZB"> Uzbekistan
	<OPTION VALUE="VUT"> Vanuatu
	<OPTION VALUE="VAT"> Vatican City State (Holy See)
	<OPTION VALUE="VEN"> Venezuela
	<OPTION VALUE="VNM"> Viet Nam
	<OPTION VALUE="VGB"> Virgin Islands (British)
	<OPTION VALUE="VIR"> Virgin Islands (U.S.)
	<OPTION VALUE="WLF"> Wallis And Futuna Islands
	<OPTION VALUE="ESH"> Western Sahara
	<OPTION VALUE="YEM"> Yemen
	<OPTION VALUE="YUG"> Yugoslavia
	<OPTION VALUE="ZAR"> Zaire
	<OPTION VALUE="ZMB"> Zambia
	<OPTION VALUE="ZWE"> Zimbabwe
	</SELECT>
<%
End Sub


Sub PrintStatesProvinces( strName )
%>
	<SELECT Name="<%=strName%>" size="1">
	<OPTION value="" >(Req'd for US/Canada)</OPTION>
	<OPTION VALUE="AL"> Alabama
	<OPTION VALUE="AK"> Alaska
	<OPTION VALUE="AZ"> Arizona
	<OPTION VALUE="AR"> Arkansas
	<OPTION VALUE="CA"> California
	<OPTION VALUE="CO"> Colorado
	<OPTION VALUE="CT"> Connecticut
	<OPTION VALUE="DE"> Delaware
	<OPTION VALUE="DC"> District of Columbia
	<OPTION VALUE="FL"> Florida
	<OPTION VALUE="GA"> Georgia
	<OPTION VALUE="HI"> Hawaii
	<OPTION VALUE="ID"> Idaho
	<OPTION VALUE="IL"> Illinois
	<OPTION VALUE="IN"> Indiana
	<OPTION VALUE="IA"> Iowa
	<OPTION VALUE="KS"> Kansas
	<OPTION VALUE="KY"> Kentucky
	<OPTION VALUE="LA"> Louisiana
	<OPTION VALUE="ME"> Maine
	<OPTION VALUE="MD"> Maryland
	<OPTION VALUE="MA"> Massachusetts
	<OPTION VALUE="MI"> Michigan
	<OPTION VALUE="MN"> Minnesota
	<OPTION VALUE="MS"> Mississippi
	<OPTION VALUE="MO"> Missouri
	<OPTION VALUE="MT"> Montana
	<OPTION VALUE="NE"> Nebraska
	<OPTION VALUE="NV"> Nevada
	<OPTION VALUE="NH"> New Hampshire
	<OPTION VALUE="NJ"> New Jersey
	<OPTION VALUE="NM"> New Mexico
	<OPTION VALUE="NY"> New York
	<OPTION VALUE="NC"> North Carolina
	<OPTION VALUE="ND"> North Dakota
	<OPTION VALUE="OH"> Ohio
	<OPTION VALUE="OK"> Oklahoma
	<OPTION VALUE="OR"> Oregon
	<OPTION VALUE="PA"> Pennsylvania
	<OPTION VALUE="RI"> Rhode Island
	<OPTION VALUE="SC"> South Carolina
	<OPTION VALUE="SD"> South Dakota
	<OPTION VALUE="TN"> Tennessee
	<OPTION VALUE="TX"> Texas
	<OPTION VALUE="UT"> Utah
	<OPTION VALUE="VT"> Vermont
	<OPTION VALUE="VA"> Virginia
	<OPTION VALUE="WA"> Washington
	<OPTION VALUE="WV"> West Virginia
	<OPTION VALUE="WI"> Wisconsin
	<OPTION VALUE="WY"> Wyoming
	<OPTION value=""> --
	<OPTION VALUE="AA"> Armed Forces the Americas
	<OPTION VALUE="AE"> Armed Forces Europe
	<OPTION VALUE="AP"> Armed Forces Pacific
	<OPTION value=""> --
	<OPTION VALUE="AB"> Alberta
	<OPTION VALUE="BC"> British Columbia
	<OPTION VALUE="MB"> Manitoba
	<OPTION VALUE="NB"> New Brunswick
	<OPTION VALUE="NF"> Newfoundland
	<OPTION VALUE="NT"> Northwest Territories
	<OPTION VALUE="NS"> Nova Scotia
	<OPTION VALUE="ON"> Ontario
	<OPTION VALUE="PE"> Prince Edward Island
	<OPTION VALUE="QC"> Quebec
	<OPTION VALUE="SK"> Saskatchewan
	<OPTION VALUE="YT"> Yukon
	</SELECT>
<%
End Sub

%>

<!-- #include file="closedsn.asp" -->

<!-- #include file="footer.asp" -->