<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Edit Customer</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

'if not LoggedStaff() then Redirect("login.asp?Source=customer_edit.asp&ID=" & intID)

strSubmit = Request("Submit")

if strSubmit = "Update" then


	'We have to verify the card
	if Request("VerifyCard") = "1" then VerifyCard

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	Query = "SELECT * FROM Customers WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	if Request("VersionOther") <> "" then
		strVersion = Request("VersionOther")
	else
		strVersion = Request("Version")
	end if
	rsUpdate("Version") = strVersion

	oldSubDir = rsUpdate("SubDirectory")
	if oldSubDir = Request("SubDirectory") then
		blSameDir = true
	else
	'We changed the sub-dir, so chenge it
		blSameDir = false
		newSubDir = Request("SubDirectory")

		'Make sure the new folder doesn't exist
		strPath = Server.MapPath("..\" & oldSubDir)

		strNewPath = Server.MapPath("..\" & newSubDir)
		if FileSystem.FolderExists( strNewPath ) then
			Set rsUpdate = Nothing
			Set FileSystem = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but that subdirectory has already been taken.  Please choose a different one."))
		end if

		rsUpdate("SubDirectory") = newSubDir
		'old path = strpath, new path = strNewPath

		'Make sure the old folder exists.  some sites (such as foreign or email hosting) may not have a sub-dir
		if FileSystem.FolderExists(strPath) then
			oldDirExists = true

			FileSystem.CreateFolder strNewPath
			FileSystem.CopyFolder strPath , strNewPath
			FileSystem.DeleteFolder strPath

			'now that we renamed it, set strpath = strNewPath
			strPath = strNewPath

		else
			oldDirExists = false
		end if
	end if


	rsUpdate("Date") = AssembleDate("Date")

	rsUpdate("Organization") = Format(Request("Organization"))
	rsUpdate("FirstName") = Format(Request("FirstName"))
	rsUpdate("LastName") = Format(Request("LastName"))
	rsUpdate("EMail") = Format(Request("EMail"))
	rsUpdate("Street1") = Format(Request("Street1"))
	rsUpdate("Street2") = Format(Request("Street2"))
	rsUpdate("City") = Format(Request("City"))
	rsUpdate("State") = Format(Request("State"))
	rsUpdate("Zip") = Format(Request("Zip"))
	rsUpdate("Country") = Format(Request("Country"))
	rsUpdate("Phone") = Format(Request("Phone"))

	rsUpdate("DomainName") = Request("DomainName")
	rsUpdate("UseDomain") = cInt(Request("UseDomain"))


	rsUpdate("FreeSite") = cInt(Request("FreeSite"))
	if Request("BillingTypeOther") <> "" then
		rsUpdate("BillingType") = Request("BillingTypeOther")
	else
		rsUpdate("BillingType") = Request("BillingType")
	end if


	rsUpdate("ChargeAdditionalFees") = cInt(Request("ChargeAdditionalFees"))

	rsUpdate("BillingCycleMonths") = cInt(Request("BillingCycleMonths"))


	rsUpdate("BillingStreet1") = Format(Request("BillingStreet1"))
	rsUpdate("BillingStreet2") = Format(Request("BillingStreet2"))
	rsUpdate("BillingCity") = Format(Request("BillingCity"))
	rsUpdate("BillingState") = Format(Request("BillingState"))
	rsUpdate("BillingZip") = Format(Request("BillingZip"))
	rsUpdate("BillingCountry") = Format(Request("BillingCountry"))
	rsUpdate("BillingPhone") = Format(Request("BillingPhone"))

	rsUpdate("CCType") = Request("CCType")
	rsUpdate("CCFirstName") = Request("CCFirstName")
	rsUpdate("CCLastName") = Request("CCLastName")
	rsUpdate("CCCompany") = Request("CCCompany")
	rsUpdate("CCNumber") = Request("CCNumber")
	rsUpdate("CCExpDate") = AssembleDate("CCExpDate")

	rsUpdate.Update
	rsUpdate.Close

	Query = "SELECT Title FROM Configuration WHERE CustomerID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
		rsUpdate("Title") = Format(Request("Title"))

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing


	strPath = Server.MapPath("..\" & Request("SubDirectory") ) & "\"
	strSource = "No"

	'Custom sub for the buttons
	Sub PrintCustomMenu
		Exit Sub
	End Sub

	'Custom sub for the buttons
	Sub PrintCustomFooter
		Exit Sub
	End Sub

	'Custom sub for the buttons
	Function GetCustomPreload
		GetCustomPreload = ""
	End Function

	CustomerID = intID

	strFileName = strPath & "\" & "write_constants.asp"
	constantFile = FileSystem.FileExists(strFileName)
	if constantFile then GetURL "http://www.Grouploop.com/" & Request("SubDirectory") & "/write_constants.asp"

	strFileName = strPath & "\" & "write_header_footer.asp"
	headerFile = FileSystem.FileExists(strFileName)
	if headerFile then GetURL "http://www.Grouploop.com/" & Request("SubDirectory") & "/write_header_footer.asp"


	'let's indicate whether or not they have a new directory
	if not blSameDir then
		Redirect("customer_edit.asp?Submit=Changed&ID=" & intID & "&OldDir=" & Server.URLEncode(oldSubDir) )
	else
		Redirect("customer_edit.asp?Submit=Changed&ID=" & intID )
	end if
	Set FileSystem = Nothing

elseif strSubmit = "Changed" then


	oldSubDir = Request("OldDir")

	if oldSubDir = "" then
%>
	<p>The customer has been updated.  <a href="customer_edit.asp?ID=<%=intID%>">Edit them again.</a><br>
	<a href="customer_view.asp?ID=<%=intID%>">View their details.</a><br>
	<a href="customers.asp">Browse through customers.</a><br>	
	</p>
<%
	'Give them the option to keep the old directory active
	elseif Request("Redirector") = "" then
%>
	<form method="post" action="customer_edit.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
		Since you changed the sub-directory, you may want to keep the old directory, and have it redirect users to the new 
		sub-directory.  This will last for over a month, and will inform current members of the change.  For sites currently 
		being used, this is <b>highly recommended.</b><br>


		<input type="button" value="Keep Old Directory Active" onClick="Redirect('customer_edit.asp?ID=<%=intID%>&Submit=Changed&OldDir=<%=oldSubDir%>&Redirector=Yes')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Forget Old Directory" onClick="Redirect('customer_edit.asp?ID=<%=intID%>&Submit=Changed')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;


<%
	'Put a redirector script in the old folder
	else
		oldSubDir = Request("OldDir")
		CreateRedirector oldSubDir, intID

		Redirect("customer_edit.asp?ID=" & intID & "&Submit=Changed")
	end if

else
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsPage.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing

	Set OwnerID = rsPage("OwnerID")
	Set SignupDate = rsPage("SignupDate")
	Set MasterID = rsPage("MasterID")
	Set ParentID = rsPage("ParentID")
	Set FreeSite = rsPage("FreeSite")

	Set UseDomain = rsPage("UseDomain")
	Set DomainName = rsPage("DomainName")
	Set SubDirectory = rsPage("SubDirectory")

	Set Title = rsPage("Title")

	Set Organization = rsPage("Organization")
	Set FirstName = rsPage("FirstName")
	Set LastName = rsPage("LastName")
	Set Street1 = rsPage("Street1")
	Set Street2 = rsPage("Street2")
	Set City = rsPage("City")
	Set State = rsPage("State")
	Set Zip = rsPage("Zip")
	Set Country = rsPage("Country")
	Set Phone = rsPage("Phone")

	Set BillingType = rsPage("BillingType")
	Set BillingStreet1 = rsPage("BillingStreet1")
	Set BillingStreet2 = rsPage("BillingStreet2")
	Set BillingCity = rsPage("BillingCity")
	Set BillingState = rsPage("BillingState")
	Set BillingZip = rsPage("BillingZip")
	Set BillingPhone = rsPage("BillingPhone")
	Set BillingCountry = rsPage("BillingCountry")
	Set CCFirstName  = rsPage("CCFirstName")
	Set CCLastName  = rsPage("CCLastName")
	Set CCCompany = rsPage("CCCompany")
	Set CCType = rsPage("CCType")
	Set CCNumber  = rsPage("CCNumber")
	Set CCExpDate   = rsPage("CCExpDate")

	Set TransID = rsPage("TransID")
	Set MerchantClientIDNumber = rsPage("MerchantClientIDNumber")
	Set MerchantBank = rsPage("MerchantBank")
	Set ChargeAdditionalFees = rsPage("ChargeAdditionalFees")
	Set BillingCycleMonths = rsPage("BillingCycleMonths")



	Set Version = rsPage("Version")

	Set MemberFirstName = rsPage("MemberFirstName")
	Set MemberLastName = rsPage("MemberLastName")
	Set HomeStreet = rsPage("HomeStreet")
	Set HomeCity = rsPage("HomeCity")
	Set HomeState = rsPage("HomeState")
	Set HomeZip = rsPage("HomeZip")
	Set HomePhone = rsPage("HomePhone")
	Set Beeper = rsPage("Beeper")
	Set CellPhone = rsPage("CellPhone")


	Set NickName = rsPage("NickName")
	Set Password = rsPage("Password")

	Set EMail = rsPage("EMail")
	Set MemberEMail1 = rsPage("EMail1")
	Set MemberEMail2 = rsPage("EMail2")


	rsPage.Filter = "ID = " & intID
%>
	* indicated required information<br>

	<script language="JavaScript">
	<!--
		//Throw out all the stuff we don't want ($)
		function ConvertDollar(currCheck) {
			if (!currCheck) return '';
			for (var i=0, currOutput='', valid="0123456789."; i<currCheck.length; i++)
				if (valid.indexOf(currCheck.charAt(i)) != -1)
					currOutput += currCheck.charAt(i);
			return currOutput;
		}


		function CardChanged(form) {
		
			var where_to= confirm('Since you changed the card number, it is recommended that you verify this card, to guarantee a genuine card.  If you wish to verify the card, click OK.  If not, click Cancel.');
			if (where_to== true) {
				form.VerifyCard[0].checked = true;
			}
			else
				{
				form.VerifyCard[1].checked = true;
			}
			return true; 
		}

		function submit_page(form) {
			//Error message variable
			var strError = "";

			//They changed the credit card number, so just double check with them about it
			if (!form.VerifyCard[0].checked){
				if (form.CCNumber.value != '<%=CCNumber%>' && form.VerifyCard[1].checked == true)
				var where_to= confirm('You have chosen to not verify the card.  Since you changed the card number, it is recommended that you verify this card.  If you wish to verify the card, click OK.  If you still do not, click Cancel.');
				if (where_to== true) {
					form.VerifyCard[0].checked = true;
				}
				else
					{
					form.VerifyCard[1].checked = true;
				}
			}			



			if (form.SubDirectory.value == "")
				strError += "          You forgot the directory. \n";
			if (form.BillingCycleMonths.value == "")
				strError += "          You forgot the billing cycle months. \n";



			if (form.Version.value == "Other" && form.VersionOther.value == "")
				strError += "          You forgot the version. \n";
			if (form.Title.value == "")
				strError += "          You forgot the site title. \n";
			if (form.FirstName.value == "" || form.LastName.value == "")
				strError += "          You forgot the name. \n";
			if (form.EMail.value == "")
				strError += "          You forgot the E-Mail address. \n";
			if (form.Street1.value == "" || form.City.value == "" || form.State.value == "" || form.Zip.value == "" || form.Country.value == "" || form.Phone.value == "" )
				strError += "          You screwed up the Contact Information. \n";
			if (form.Phone.value == "")
				strError += "          You forgot the phone number. \n";
			if (form.BillingType.value == "Other" && form.BillingTypeOther.value == "")
				strError += "          You forgot the default billing type. \n";
			if (form.BillingStreet1.value == "" || form.BillingCity.value == "" || form.BillingState.value == "" || form.BillingZip.value == "" ||  form.BillingCountry.value == "" || form.BillingPhone.value == "" )
				strError += "          You screwed up the Billing Address. \n";


			if(strError == "") {
				return true;
			}
			else{
				strError = "We have detected the following possible missing information: \n" + strError +  "\n\nIf you wish to proceed, click OK.  To correct the missing info, click Cancel.";
				var where_to= confirm(strError);
				
				if (where_to== true) {
					return true;
				}
				else
					{
					return false;
				}
			}   
		}

		function CopyAddress(form) {
			form.elements['BillingStreet1'].value = form.elements['Street1'].value;
			form.elements['BillingStreet2'].value = form.elements['Street2'].value;
			form.elements['BillingCity'].value = form.elements['City'].value;
			form.elements['BillingState'].value = form.elements['State'].value;
			form.elements['BillingZip'].value = form.elements['Zip'].value;
			form.elements['BillingCountry'].value = form.elements['Country'].value;
			form.elements['BillingPhone'].value = form.elements['Phone'].value;
			return true; 
		}

	//-->
	</SCRIPT>

<%'	
'<form method="post" action="customer_edit.asp" name="MyForm">
%>
<form method="post" action="customer_edit.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=intID%>">
		<div ID="infoParent" NAME="infoParent" CLASS=parent>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('info'); return false" ID="infoIm">
		Site Information</a>
		</td></tr></table>
		</div>
		<div ID="infoChild" NAME="infoChild" CLASS=child>

		<% PrintTableHeader 100 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Version</td>
      		<td class="<% PrintTDMain %>">
				<select name="Version" 	onChange="if (this.form.Version.value == 'Other') this.form.VersionOther.focus();" >
<%
					WriteOption "Free", "Free", Version
					WriteOption "Gold", "Gold", Version
					WriteOption "Parent", "Parent", Version
					WriteOption "Child", "Child", Version

					if Version <> "Free" and Version <> "Gold" and Version <> "Parent" and Version <> "Child" then
						WriteOption "Other", "Other", "Other"
						strOther = Version
					else
						WriteOption "Other", "Other", Version
						strOther = ""
					end if
%>
					</select>	Other <input type="text" name="VersionOther" size="15" value="<%=strOther%>">
     		</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" align="right">Sub-Directory</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="SubDirectory" size="55" value="<%=SubDirectory%>">
     		</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" align="right">Use a domain name?</td>
      		<td class="<% PrintTDMain %>"> 
       			<% PrintRadio UseDomain, "UseDomain" %> &nbsp; &nbsp; &nbsp;Domain name: <input type="text" name="DomainName" size="55" value="<%=DomainName%>">
     		</td>

		<tr>
      		<td class="<% PrintTDMain %>" align="right">* Date Created</td>
      		<td class="<% PrintTDMain %>"> 
        		<% DatePulldown "Date", SignupDate, 1 %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Organization</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Organization" size="55" value="<%=Organization%>">
     		</td>
		</tr>
		</table>


		<div ID="contactParent" NAME="contactParent" CLASS=parent>
		<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('contact'); return false" ID="contactIm">
		Contact Information</a>
		</td></tr></table>
		</div>
		<div ID="contactChild" NAME="contactChild" CLASS=child>
		<% PrintTableHeader 100 %>

		<tr>
      		<td class="<% PrintTDMain %>" align="right">* Site Title</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Title" size="55" value="<%=Title%>">
     		</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" align="right">* Owner Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="FirstName" size="40" value="<%=FirstName%>">&nbsp;&nbsp;<input type="text" name="LastName" size="40" value="<%=LastName%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="55" value="<%=EMail%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Street1" size="55" value="<%=Street1%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Address 2nd Line</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Street2" size="55" value="<%=Street2%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* City</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="City" size="55" value="<%=City%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* State</td>
      		<td class="<% PrintTDMain %>"> 
				<% PrintStates "State", State %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Zip Code</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Zip" size="10" value="<%=Zip%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Country</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Country" size="10" value="<%=Country%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Phone Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Phone" size="15" value="<%=Phone%>">
     		</td>
		</tr>
		</table>
		</div>	

		</div>
		<br>



	<div ID="billingParent" NAME="billingParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('billing'); return false" ID="billingIm">
		Billing Information</a>
		</td></tr>
	</table>
	</div>
	<div ID="billingChild" NAME="billingChild" CLASS=child>

	<% PrintTableHeader 100 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Free Site?</td>
      		<td class="<% PrintTDMain %>"> 
       			<% PrintRadio FreeSite, "FreeSite" %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Default Billing Method</td>
      		<td class="<% PrintTDMain %>"> 
				<select name="BillingType" 	onChange="if (this.form.BillingType.value == 'Other') this.form.BillingTypeOther.focus();" >
<%
					WriteOption "CreditCard", "Credit Card", BillingType
					WriteOption "Check", "Check", BillingType
					WriteOption "Other", "Other", BillingType

					if BillingType <> "CreditCard" and BillingType <> "Check" then
						strOther = BillingType
					else
						strOther = ""
					end if
%>
					</select>	Other <input type="text" name="BillingTypeOther" size="15" value="<%=strOther%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Length, in months of their Billing Cycle</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingCycleMonths" size="5" value="<%=BillingCycleMonths%>">
     		</td>
		</tr>

		<tr> 
      		<td class="<% PrintTDMain %>" align="right">When charging customer, should additional fees be automatically calculated?</td>
      		<td class="<% PrintTDMain %>"> 
       			<% PrintRadio ChargeAdditionalFees, "ChargeAdditionalFees" %>
     		</td>
		</tr>



	</table>
	<div ID="billingaddParent" NAME="billingaddParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('billingadd'); return false" ID="billingaddIm">
		Billing Address</a>
		</td></tr>
	</table>
	</div>
	<div ID="creditParent" NAME="creditParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('credit'); return false" ID="creditIm">
		Credit Card Information</a>
		</td></tr>
	</table>
	</div>
	<div ID="billingaddChild" NAME="billingaddChild" CLASS=child>

	<% PrintTableHeader 100 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
				Billing Address
     		</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" colspan=2 align="center"> 
				<input type="button" value="Use Contact Address" onClick="CopyAddress(this.form);">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingStreet1" size="55" value="<%=BillingStreet1%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Address 2nd Line</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingStreet2" size="55" value="<%=BillingStreet2%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* City</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingCity" size="55" value="<%=BillingCity%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* State</td>
      		<td class="<% PrintTDMain %>"> 
				<% PrintStates "BillingState", State %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Zip Code</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingZip" size="10" value="<%=BillingZip%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Country</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingCountry" size="10" value="<%=BillingCountry%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Phone Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="BillingPhone" size="15" value="<%=BillingPhone%>">
     		</td>
		</tr>
		</table>
	</div>



	<div ID="creditChild" NAME="creditChild" CLASS=child>
	<% PrintTableHeader 100 %>


		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Credit Card Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Verify Card Number?</td>
      		<td class="<% PrintTDMain %>"> 
       			<% PrintRadio 0, "VerifyCard" %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Credit Card Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCNumber" size="20" value="<%=CCNumber%>" onChange="CardChanged(this.form);">
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" align="right">Card Type</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<select name="CCType" size=1>
<%
					WriteOption "VISA", "VISA", CCType
					WriteOption "MasterCard", "MasterCard", CCType
					WriteOption "AmEx", "American Express", CCType
%>
				</select>
			</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" align="right">Name on Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCFirstName" size="40" value="<%=CCFirstName%>">&nbsp;&nbsp;<input type="text" name="CCLastName" size="40" value="<%=CCLastName%>">
     		</td>
		</tr>
		<tr>
      		<td class="<% PrintTDMain %>" align="right">Company on Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCompany" size="40" value="<%=CCCompany%>">
     		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Expiration Date
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
       		<% DatePulldown "CCExpDate", CCExpDate, 0 %>

			</td>
		</tr>
	</table>

	</div>

	</div>


<br>
	<% PrintTableHeader 100 %>
		<tr>
    		<td colspan="2" align="center" class="TDBlank">

				<input type="submit" name="Submit" value="Update">
	   		</td>
		</tr>

	</table>
</form>
<%
	rsPage.Close
	Set rsPage = Nothing
end if



Sub	CreateRedirector( oldSubDir, intID )

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	strNewPath = Server.MapPath("..\" & oldSubDir) & "\"

	'This should never happen.  We rename the sub-dir, so it should be free
	if FileSystem.FolderExists(strNewPath) then
			Redirect("message.asp?Message=" & Server.URLEncode("For some crazy reason, the old directory is still there.  Therefore, the redirector could not be set up.  <a href='http://www.GroupLoop.com/" & oldSubDir & "'>See what's there right now.</a>"))
	end if

	FileSystem.CreateFolder(strNewPath)


	'Get the subdirectory
	Set Command = Server.CreateObject("ADODB.Command")
	With Command
		'Check the scheme to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandType = adCmdStoredProc

		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intID
		.Execute , , adExecuteNoRecords
		strSubDir = .Parameters("@SubDirectory")
		strTitle = .Parameters("@Title")

		if strSubDir = "" then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The new directory is not in the database for some reason"))
		end if

	End With
	Set Command = Nothing


	strQuote = Chr(34)

	PageSource = "<html>"  & vbCrlf & _
		"<!--RedirectUser-->"  & vbCrlf & _
		"<!--" & FormatDateTime(Date,2) & "*" & oldSubDir & "*" & strSubDir & "-->"  & vbCrlf & _
		"<!--Format: DateCreated*OldDir(this dir)*NewDir-->"  & vbCrlf & _

		"<head>"  & vbCrlf & _
		"<title>Site Address Change!</title>" & vbCrlf & _
		"<meta http-equiv=" & strQuote & "Content-Type" & strQuote & " content=" & strQuote & "text/html; charset=iso-8859-1" & strQuote & ">"  & vbCrlf & _
		"<meta http-equiv=" & strQuote & "refresh" & strQuote & " content=" & strQuote & "35;URL=http://www.GroupLoop.com/newsub" & strQuote & ">"  & vbCrlf & _
		"</head>"  & vbCrlf & _

		"<body bgcolor=" & strQuote & "#FFFFFF" & strQuote & " text=" & strQuote & "#000000" & strQuote & ">"  & vbCrlf & _
		"<Script Language = " & strQuote & "JavaScript" & strQuote & " Type=" & strQuote & "Text/JavaScript" & strQuote & ">"  & vbCrlf & _
		"<!--  Hide script from older browsers"  & vbCrlf & _
		"var urlAddress = " & strQuote & "http://www.GroupLoop.com/" & strSubDir & strQuote & ";"  & vbCrlf & _
		"var pageName = " & strQuote & Format(strTitle) & strQuote & ";"  & vbCrlf & _

		"function addToFavorites()"  & vbCrlf & _
		"{"  & vbCrlf & _
		"	if (window.external)"  & vbCrlf & _
		"	{"  & vbCrlf & _
		"		window.external.AddFavorite(urlAddress,pageName)"  & vbCrlf & _
		"	}"  & vbCrlf & _
		"	else"  & vbCrlf & _
		"	{ "  & vbCrlf & _
		"		alert(" & strQuote & "Sorry! Your browser doesn't support this function." & strQuote & ");"  & vbCrlf & _
		"	}"  & vbCrlf & _
		"}"  & vbCrlf & _
		"// -->"  & vbCrlf & _
		"</script>"  & vbCrlf & _
		"<p align=" & strQuote & "center" & strQuote & "><img src=" & strQuote & "http://www.GroupLoop.com/homegroup/images/title.gif" & strQuote & "></p>"  & vbCrlf & _
		"<p align=" & strQuote & "center" & strQuote & "><i><font size=" & strQuote & "+2" & strQuote & ">" & strTitle & "</font></i></p>"  & vbCrlf & _
		"<p><b><font color=" & strQuote & "#FF0000" & strQuote & " size=" & strQuote & "+1" & strQuote & ">The address of this web page has changed.</font></b></p>"  & vbCrlf & _
		"<p align=" & strQuote & "left" & strQuote & "><font size=" & strQuote & "+3" & strQuote & ">The new address is <a href=" & strQuote & "http://www.GroupLoop.com/" & strSubDir & strQuote & ">http://www.GroupLoop.com/" & strSubDir & "</a></font></p>"  & vbCrlf & _
		"<p><i>You will automatically be directed to the new site in 30 seconds.<br>" & vbCrlf & _
		"  However, this reminder is only temporary, so<br>"  & vbCrlf & _
		"  </i><b><font size=" & strQuote & "+2" & strQuote & "><a href=javascript:addToFavorites()>bookmark the new site.</a></font></b></p>"  & vbCrlf & _
		"</body>"  & vbCrlf & _
		"</html>"

	Set ConstFile = FileSystem.CreateTextFile(strNewPath & "index.htm")


	ConstFile.WriteLine PageSource
	ConstFile.Close
	Set ConstFile = Nothing

	Set FileSystem = Nothing

End Sub


Sub VerifyCard
	'Verify the Credit Card HERE

	' Create xAuthorize object.
	Dim objAuthorize
	Set objAuthorize = Server.CreateObject("xAuthorize.Process")
	' Initialize xAuthorize for a new transaction.
	objAuthorize.Initialize

	' Set object properties.
	objAuthorize.Processor = "AUTHORIZE_NET"

	objAuthorize.FirstName = Request("CCFirstName")
	objAuthorize.LastName = Request("CCLastName")
	objAuthorize.Company = Request("CCCompany")
	objAuthorize.Address = Request("BillingStreet1") & " " & Request("BillingStreet2")
	objAuthorize.City = Request("BillingCity")
	objAuthorize.State = Request("BillingState")
	objAuthorize.Zip = Request("BillingZip")
	objAuthorize.Country = Request("BillingCountry")

	objAuthorize.CustomerID = intID
	objAuthorize.InvoiceNumber = intID

	objAuthorize.Login = "OurPage"
	objAuthorize.Password = "hgf554jh"

	objAuthorize.CardNumber = Request("CCNumber")
	objAuthorize.CardType = Request("CCType")
	CCExpDate = AssembleDate("CCExpDate")
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

	if strError <> "" then Redirect("message.asp?Message=" & Server.URLEncode(strError))
End Sub
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->