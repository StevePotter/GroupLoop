<!-- #include file="header.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="dsn.asp" -->
<% AddHit "pricing.asp" %>
<!-- #include file="closedsn.asp" -->


<p class="Heading" align=center>Pricing</p>
<b>There are three versions of GroupLoop sites: the Free Version, Gold Version, and Multi-Site Version.</b>

<table cellspacing=2 cellpadding=2 border=0>
	<tr>
		<td align=center class="TDHeader">&nbsp;</td>
		<td align=center class="TDHeader">Free Version</td>
		<td align=center class="TDHeader">Gold Version</td>
		<td align=center class="TDHeader">Multi-Site Version</td>
	</tr>
	<tr>
		<td align=left class="TDHeader">Description</td>
		<td align=left class="<%PrintTDMain%>">Perfect way for a group to see what GroupLoop has to offer without a commitment.</td>
		<td align=left class="<%PrintTDMain%>">What most customers use.  Includes all our features, with no limitations.  Perfect for any group!</td>
		<td align=left class="<%PrintTDMainSwitch%>">Unleash the true power of GroupLoop with a community of linked web sites. <a href="multisite.asp">Details.</a></td>
	</tr>
	<tr>
		<td align=left class="TDHeader">Monthly Fee</td>
		<td align=left class="<%PrintTDMain%>">Totally Free</td>
		<td align=left class="<%PrintTDMain%>">* $20 per month</td>
		<td align=left class="<%PrintTDMainSwitch%>"><a href="multisite.asp">Click here for fees</a></td>
	</tr>
	<tr>
		<td align=left class="TDHeader">Features</td>
		<td align=left class="<%PrintTDMain%>" valign="top">
		<ul>
			<li>All the sections (announcements, calendar, photos, etc.)</li>
			<li>Banner advertisements.</li>
			<li><b>Expires after 2 months.</b></li>
			<li>Up to 5 members.</li>
			<li>5 Megabytes of photo space (about 80 photos).</li>
			<li>10 Megabytes of media space for your favorite sounds, movies, songs, etc.</li>
			<li><i>Upgrade at any time.</i></li>
		</ul>
			<div align="center">
			<form METHOD="post" ACTION="signup2.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
			<input type="submit" name="Submit" value="Use Free Version">
			</form>
			</div>
		
		</td>
		<td align=left class="<%PrintTDMain%>" valign="top">
			<ul>
				<li><a href="advancedfeatures.asp">All GroupLoop Features</a></li>
				<li>No banner advertisements whatsoever.</li>
				<li>Your site hosted on a secure, high-speed server, with %99.9 uptime, guaranteed.</li>
				<li>All the sections (announcements, calendar, photos, etc.)</li>
				<li>30 Megabytes of photo space (about 500 photos), with the option to add more.</li>
				<li>40 Megabytes of media space for your favorite sounds, movies, songs, etc, with the option to add more.</li>
				<li>Unlimited number of members*.</li>
				<li>The option to have your own domain name (www.YourGroup.com).</li>
				<li>Free use of a secure connection to keep vital member information safe.</li>
				<li>Option to have additional custom sections created.</li>
				<li>Extensive customer support from our caring, helpful experts.</li>
			</ul>
			<div align="center">
			<form METHOD="post" ACTION="signup2.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
			<input type="submit" name="Submit" value="Use Gold Version">
			</form>
			</div>
		</td>
		<td align=left class="<%PrintTDMainSwitch%>" valign="top">
			<ul>
				<li><a href="advancedfeatures.asp">All GroupLoop Features</a></li>
				<li>Perfect for organizations requiring more than one site (religious organziations, sport leagues, and any sub-divided organization).</li>
				<li>Unlimited number of separate, linked sites.</li>
				<li>Top priority with customer support.</li>
				<li>Custom payment plans available (e-mail <a href="mailto:accounts@grouploop.com">accounts@grouploop.com</a>).</li>
				<li>Each site has all the Gold version features.</li>
				<li>Members of more than one site can add items to multiple sites at once.</li>
				<li>Ability to add and remove sites at will.</li>
			</ul>
			<div align="center">
			<form METHOD="post" ACTION="signup1a.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
			<input type="submit" name="Submit" value="Use Multi-Site Version">
			</form>
			</div>
		</td>
	</tr>
</table>


<%
'<p align="center" class=SubHeading>Why isn't everything free?</p>
'<p>
'We know there are some other free services out there for groups of people.  The problem with free 
'services is that realistically, noone offering something free can really take care of their customers. 
'<a href="why.asp#Others">Click here</a> to compare us to a few other services.
'</p>
%>


<p class="SubHeading" align=center>* Additional Fees (monthly)</p>
<table cellspacing=2 cellpadding=4 align=center>
	<tr>
		<td align=right class="<%PrintTDMain%>">Over 20 members.</td>
		<td align=left class="<%PrintTDMainSwitch%>">$0.75 per each extra member.</td>
	</tr>
	<tr>
		<td align=right class="<%PrintTDMain%>">Domain Name (www.yourownname.com).</td>
		<td align=left class="<%PrintTDMainSwitch%>">$50 <b>one-time</b> setup fee plus $5 monthly server charge.</td>
	</tr>
	<tr>
		<td align=right class="<%PrintTDMain%>">Additional Photos Space.</td>
		<td align=left class="<%PrintTDMainSwitch%>">$0.50 for each additional megabyte.</td>
	</tr>
	<tr>
		<td align=right class="<%PrintTDMain%>">Media Section</td>
		<td align=left class="<%PrintTDMainSwitch%>">$0.50 for each additional megabyte.</td>
	</tr>
</table>


<p class="SubHeading" align=center>Custom Designing</p>
<div align="center">
	<a href="custom.asp">Click here for our custom solutions!</a>
</div>

<%
'<p class="Heading" align=center>Referral Policy!</p>
'We really want everyone to spread the word about GroupLoop.com.  So, as an incentive to do so, we've created our Referral 
'Progam.  If you have a site and get someone to list you as a reference, you get a <b>free</b> month (additional fees 
'do not apply, just the $20 base fee).  So <b><a href="college.asp">poor college kids</a></b> and anyone looking for a free site, sign up and start 
'telling your friends!
%>


<!-- #include file="footer.asp" -->