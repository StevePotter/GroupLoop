<!-- #include file="header.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="dsn.asp" -->
<% AddHit "signup1.asp" %>
<!-- #include file="closedsn.asp" -->

<p class=Heading align=center>
Step 1. Choose Your Version
</p>




<table width="100%" cellspacing=2 cellpadding=1 border=0>

	<tr>
		<td class=TDHeader align=center>
			Free Version
		</td>
		<td class=TDHeader align=center>
			Gold Version
		</td>
		<td class=TDHeader align=center>
			<a href="multisite.asp">Multi-Site Version</a>
		</td>
	</tr>
	<tr>
		<td class="BodyText" valign="top" align="left">
<br>
		- Totally Free <br>
		- All the Gold features.<br>
		- Expires after 1 month.<br>
		- Upgrade at any time. <br>
		<br>
			<form METHOD="post" ACTION="signup2.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
			if Request("ParentID") <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=Request("ParentID")%>">
<%
			end if
%>
			<input type="submit" name="Submit" value="Use Free Version">
			</form>
		</td>
		<td class="BodyText" valign="top" align="left">
<br>
				- $40 per month<br>
				- <a href="advancedfeatures.asp">All GroupLoop Features</a><br>
				- No banner advertisements whatsoever.<br>
				- Your site hosted on a secure, high-speed server, with 99.9% uptime, guaranteed.<br>
				- All the sections (announcements, calendar, photos, etc.)<br>
				- 30 Megabytes of photo space (about 500 photos), with the option to add more.<br>
				- 40 Megabytes of media space for your favorite sounds, movies, songs, etc, with the option to add more.<br>
				- Unlimited number of members.<br>
				- The option to have your own domain name (www.YourGroup.com).<br>
				- Option to have additional custom sections created.<br>
				- Extensive customer support from our caring, helpful experts.<br>
				- Additional fees apply *<br>
		<br>
			<form METHOD="post" ACTION="signup2.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
			if Request("ParentID") <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=Request("ParentID")%>">
<%
			end if
%>
			<input type="submit" name="Submit" value="Use Gold Version">
			</form>
		</td>
		<td class="BodyText" valign="top" align="left">
<br>
				- Perfect for organizations requiring more than one site (religious organziations, sport leagues, and any sub-divided organization).<br>
				- Unlimited number of separate, linked sites.<br>
				- Top priority with customer support.<br>
				- Custom payment plans available (e-mail <a href="mailto:accounts@grouploop.com">accounts@grouploop.com</a>).<br>
				- Each site has all Gold version features.<br>
				- Members of more than one site can add items to multiple sites at once.<br>
				- Ability to add and remove sites at will.<br>
		<br>

			<form METHOD="post" ACTION="signup1a.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
			if Request("ParentID") <> "" then
%>
			<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
			<input type="hidden" name="ParentID" value="<%=Request("ParentID")%>">
<%
			end if
%>
			<input type="submit" name="Submit" value="Use Multi-Site Version">
			</form>
		</td>
	</tr>

</table>

<p align=center>

</p>

* Additional fees may apply. View the <a href="pricing.asp">pricing page</a> for explanation. 

<!-- #include file="footer.asp" -->