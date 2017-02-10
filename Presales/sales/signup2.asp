<!-- #include file="..\homegroup\dsn.asp" -->
<%
if Request("ReferralID") <> "" then
	Set Command = Server.CreateObject("ADODB.Command")
	With Command
		'Check the salesman
		intReferralID = CInt(Request("ReferralID"))
		.ActiveConnection = Connect
		.CommandText = "EmployeeIDExists"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@EmployeeID") = intReferralID
		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
		'They don't exist! Get the FUCK out
		if not blExists then
			Set Command = Nothing
			Redirect("message.asp?Source=noback&Message=" & Server.URLEncode("Please close then re-open your browser and come back to GroupLoop.com.  This error is due to an invalid referral ID number.  The salesman who referred you may have cancelled their membership.  It's fine, you just need to re-open your browser."))
		end if
	End With
	Set Command = Nothing

	strOnLoad = "displayreferral()"
end if
%>

<!-- #include file="header.asp" -->

<!-- #include file="..\sourcegroup\functions.asp" -->

<p class=Heading align=center>
Step 2. Enter Your Information
</p>
<p>We need the proper information from you so we can contact you to ask questions or send you your check.  Your information  
is strictly confidential, and will not be shared with anyone.</p>

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

<%
	if Request("ReferralID") <> "" then
		Query = "SELECT FirstName, LastName FROM Employees WHERE ID = " & Request("ReferralID")
		Set rsMember = Server.CreateObject("ADODB.Recordset")
		rsMember.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		SalesmanName = Format(rsMember("FirstName") & " " & rsMember("LastName"))
%>	
		function displayreferral() {
			alert ("We detected that you have been referred to us by <%=SalesmanName%>.  There is no need to enter their ID number, as it is done automatically.  Just click OK and enter your information.  Thank you!");
		}
<%
	end if
%>

	function submit_page(form) {
		//Error message variable
		var strError = "";
<%
if Request("ReferralID") = "" then
%>

		form.Referral.value = ConvertInt(form.Referral.value);
<%
	end if
%>
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
        if(!form.Agree.checked)
			strError += "          You forgot to check the Agree box at the bottom. \n";

		if(strError == "") {
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



<form METHOD="POST" ACTION="signup3.asp" name="Signup" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<table cellspacing=1 cellpadding=2 border=0>
	<tr>
		<td class=TDHeader align=center colspan=2>
			Name and Such
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
			* Your NickName For Logging In
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
			* Password For Logging In
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
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Birthdate
		</td>
		<td class="<% PrintTDMain %>" align="left">
       		<% DatePulldown "BirthDate", "1/1/1950", 0 %>
		</td>
	</tr>

	<tr>
		<td class=TDHeader align=center colspan=2>
			Home Address
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Street Address
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
			* Country
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<%PrintCountry "Country"%>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			* Phone Number (xxx.xxx.xxxx)
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Phone" size="12">
		</td>
	</tr>
<%
if Request("ReferralID") = "" then
%>

	<tr>
		<td class=TDHeader align=center colspan=2>
			Referral Information
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align=center colspan=2>
			If you heard about us through another salesman, please enter the salesman number they gave you below.  Do not 
			enter their name.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Salesman ID Number
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<input type="text" name="Referral" size="4" value="<%=Request("ReferralID")%>">
		</td>
	</tr>
<%
else
%>
			<input type="hidden" name="Referral" value="<%=Request("ReferralID")%>">
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
			<input type="submit" name="Submit" value="Sign Me Up" >
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

<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->