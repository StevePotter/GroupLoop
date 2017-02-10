<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin() and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin() then Redirect("members.asp?Source=admin_account_upgrade.asp")
if Version <> "Free" and Request("Submit") <> "Changed" then Redirect("error.asp")
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Upgrade To Gold Version</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
strSubmit = Request("Submit")

if strSubmit = "Upgrade" then
	if Request("Street1") = "" or Request("City") = "" or _
	( ( Request("Country") = "USA" or Request("Country") = "CAN" ) AND Request("State") = "" ) or Request("Zip") = "" or _
	Request("CCNumber") = "" then Redirect("incomplete.asp")

	strCCFirstName = Request("CCFirstName")
	strCCLastName = Request("CCLastName")
	strCCCompany = Request("CCCompany")
	strCCType = Request("CCType")
	strCCNumber = Request("CCNumber")
	intCCExpMonth = CInt(Request("CCExpMonth"))
	intCCExpYear = CInt(Request("CCExpYear"))
	CCExpDate = CDate(intCCExpMonth & "/01/" & intCCExpYear)
	strStreet1 = Request("Street1")
	strStreet2 = Request("Street2")
	strCity = Request("City")
	strState = Request("State")
	strZip = Request("Zip")
	strCountry = Request("Country")


	'Verify the Credit Card HERE

	' Create xAuthorize object.
	Dim objAuthorize
	Set objAuthorize = Server.CreateObject("xAuthorize.Process")
	' Initialize xAuthorize for a new transaction.
	objAuthorize.Initialize

	' Set object properties.
	objAuthorize.Processor = "AUTHORIZE_NET"

	objAuthorize.FirstName = strCCFirstName
	objAuthorize.LastName = strCCLastName
	objAuthorize.Company = strCCCompany
	objAuthorize.Address = strStreet1 & " " & strStreet2
	objAuthorize.City = strCity
	objAuthorize.State = strState
	objAuthorize.Zip = strZip
	objAuthorize.Country = strCountry

	objAuthorize.CustomerID = CustomerID
	objAuthorize.InvoiceNumber = CustomerID

	objAuthorize.Login = "OurPage"
	objAuthorize.Password = "hgf554jh"

	objAuthorize.CardNumber = strCCNumber
	objAuthorize.CardType = strCCType
	objAuthorize.ExpDate = Month(CCExpDate) & "/" & Year(CCExpDate)

	objAuthorize.Amount = 20
	objAuthorize.TransType = "AUTH_ONLY"
	objAuthorize.Description = "GroupLoop.com monthly charge."

	objAuthorize.EmailMerchant = false
	objAuthorize.EmailCustomer = false

	' Initiate transation processing
	objAuthorize.Process

	strTransID = objAuthorize.TransID

	If objAuthorize.ErrorCode = 0 Then
		 ' Communication was successful.
		 ' Examine Results
		 If objAuthorize.ResponseCode <> 1 then
			strError = "Sorry, but there is the following problem with the card you entered:<br><font size='+1'>" & objAuthorize.ResponseReasonText & "</font><br><br>Make sure you have the correct address, card number, and expiration date."
		end if
	Else
		Select Case objAuthorize.ErrorCode
			Case -1
				strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
			Case -2
				strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
			Case Else
				strError = "Sorry, an unknown error occured with the authorization network.  Please try again, and notify support@keist.com if this keeps happening."
		End Select
	End If

	Set objAuthorize = Nothing


	'We had a problem
	if strError <> "" then
		Redirect("message.asp?Message=" & Server.URLEncode(strError) )
	end if

	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'See if they need the media section
		.ActiveConnection = Connect
		.CommandText = "UpgradeSite"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@CCFirstName") = strCCFirstName
		.Parameters("@CCLastName") = strCCLastName
		.Parameters("@CCCompany") = strCCCompany
		.Parameters("@CCType") = strCCType
		.Parameters("@CCNumber") = strCCNumber
		.Parameters("@CCExpdate") = CCExpDate
		.Parameters("@TransID") = strTransID
		.Parameters("@BillingStreet1") = strStreet1
		if strStreet2 <> "" then .Parameters("@BillingStreet2") = strStreet2
		if strCity <> "" then .Parameters("@BillingCity") = strCity
		if strState <> "" then .Parameters("@BillingState") = strState
		if strZip <> "" then .Parameters("@BillingZip") = strZip
		if strCountry <> "" then .Parameters("@BillingCountry") = strCountry

		.Execute , , adExecuteNoRecords
	End With

	Set Command = Nothing

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"

	strSubject = "Customer has upgraded to Gold Version"

	strBody = Title & "<br>" & _
		"<a href=" & NonSecurePath & ">" & NonSecurePath & "</a><br>" & _
		"CustomerID: " & CustomerID & "<br>" & _
		"Transaction ID: " & strTransID & "<br>"

	Mailer.Subject = strSubject
	Mailer.BodyText = strBody

	Mailer.AddRecipient "GroupLoop Accounts", "accounts@grouploop.com"
	Mailer.SendMail

	Set Mailer = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_constants.asp" -->
<%
	Redirect("write_header_footer.asp?Source=message.asp?Message=" & Server.URLEncode("NOBACKYour site has been upgraded to the Gold Version.  We hope you enjoy your site as much as possible.  Thank you so much!") )
else
%>
	<script language="JavaScript">
		function submit_page(form) {
			//Error message variable
			var strError = "";

			if(!form.Agree.checked)
				strError += "You forgot to check the authorize box.\n";
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
			if(form.Street1.value == "")
				strError += "          You forgot your Street Address. \n";
			if(form.City.value == "")
				strError += "          You forgot your City. \n";
			if(form.Zip.value == "")
				strError += "          You forgot your Zip Code. \n";
			if(form.State.value == "" && (form.Country.value == "USA" || form.Country.value == "CAN"))
				strError += "          You forgot your State. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can sign up: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	</SCRIPT>

	<p>We thank you for deciding to upgrade to the Gold Version.  GroupLoop.com was founded on a dream, and 
	you are helping make that dream possible.</p>
	<p><b>This is a totally secure connection, and you should have no concerns about the safety of your credit card number.  
	Security is our highest concern, and your vital information is in good hands.</b></p>
	<form METHOD="POST" ACTION="https://www.OurClubPage.com/<%=SubDirectory%>/admin_account_upgrade.asp" name="Signup" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
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
			* Street Address For Card
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
			* City
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
			* Zip Code
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Zip" size="8">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Country
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<%PrintCountry "Country"%>
		</td>
	</tr>
	<tr>
		<td class=<% PrintTDMain %> align=center colspan=2>
			* I verify that I am the person who owns the above credit card, and I authorize charges from GroupLoop.com <input type="checkbox" name="Agree" value="Yes">
		</td>
	</tr>
	<tr>
		<td class=<% PrintTDMain %> align=center colspan=2>
			<input type="submit" name="Submit" value="Upgrade" >
		</td>
	</tr>
  	</table>
	</form>
<%
end if
%>

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