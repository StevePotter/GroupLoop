<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=customers.asp")
%>

<p class=Heading align=center>Customers</p>

<form METHOD="POST" ACTION="customers.asp">
	View 
	<select name="Version" onChange="this.form.submit();">
<%
strVersion = Request("Version")
	WriteOption "All", "All", strVersion
	WriteOption "Free", "Free", strVersion
	WriteOption "Gold", "Gold", strVersion
	WriteOption "Parent", "Parent", strVersion
	WriteOption "Child", "Child", strVersion
	WriteOption "Other", "Other", strVersion
%>
	</select>	
	Customers &nbsp;&nbsp;&nbsp;&nbsp;
	Or Search For <input type="text" name="Keywords" size="25">
	<input type="submit" name="Submit" value="Go"><br>
</form>
<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")


'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then

	Set rsList = Server.CreateObject("ADODB.Recordset")
	rsList.CacheSize = 100

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsList.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing

		Set ID = rsList("ID")
		Set OwnerID = rsList("OwnerID")
		Set SignupDate = rsList("SignupDate")
		Set MasterID = rsList("MasterID")
		Set ParentID = rsList("ParentID")

		Set DomainName = rsList("DomainName")
		Set SubDirectory = rsList("SubDirectory")

		Set Title = rsList("Title")

		Set Organization = rsList("Organization")
		Set FirstName = rsList("FirstName")
		Set LastName = rsList("LastName")
		Set Street1 = rsList("Street1")
		Set Street2 = rsList("Street2")
		Set City = rsList("City")
		Set State = rsList("State")
		Set Zip = rsList("Zip")
		Set Country = rsList("Country")
		Set Phone = rsList("Phone")

		Set BillingType = rsList("BillingType")
		Set BillingStreet1 = rsList("BillingStreet1")
		Set BillingStreet2 = rsList("BillingStreet2")
		Set BillingCity = rsList("BillingCity")
		Set BillingState = rsList("BillingState")
		Set BillingZip = rsList("BillingZip")
		Set BillingPhone = rsList("BillingPhone")
		Set BillingCountry = rsList("BillingCountry")
		Set CCCompany = rsList("CCCompany")
		Set CCType = rsList("CCType")
		Set TransID = rsList("TransID")
		Set MerchantClientIDNumber = rsList("MerchantClientIDNumber")
		Set MerchantBank = rsList("MerchantBank")

		Set MemberFirstName = rsList("MemberFirstName")
		Set MemberLastName = rsList("MemberLastName")
		Set HomeStreet = rsList("HomeStreet")
		Set HomeCity = rsList("HomeCity")
		Set HomeState = rsList("HomeState")
		Set HomeZip = rsList("HomeZip")
		Set HomePhone = rsList("HomePhone")
		Set Beeper = rsList("Beeper")
		Set CellPhone = rsList("CellPhone")


		Set NickName = rsList("NickName")
		Set Password = rsList("Password")

		Set EMail = rsList("EMail")
		Set MemberEMail1 = rsList("EMail1")
		Set MemberEMail2 = rsList("EMail2")


	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if

if intSearchID <> "" then
	'Their search came up empty
	if intSearchID = 0 then
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="customers.asp">click here</a> to view all customers.</p>
<%
'-----------------------Begin Code----------------------------
	else
		'They have search results, so lets list their results
		Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
		rsPage.CacheSize = PageSize
		Set TargetID = rsPage("TargetID")
%>
		<form METHOD="POST" ACTION="customers.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle

		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100

		Set cmdTemp = Server.CreateObject("ADODB.Command")
		cmdTemp.ActiveConnection = Connect
		cmdTemp.CommandText = "GetSiteInfoRecordSet"
		cmdTemp.CommandType = adCmdStoredProc

		rsList.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set cmdTemp = Nothing

		Set ID = rsList("ID")
		Set CustomerID = rsList("ID")
		Set OwnerID = rsList("OwnerID")
		Set SignupDate = rsList("SignupDate")

		Set UseDomain = rsList("UseDomain")
		Set DomainName = rsList("DomainName")
		Set SubDirectory = rsList("SubDirectory")


		Set Title = rsList("Title")
		Set NickName = rsList("NickName")
		Set Password = rsList("Password")
		Set FirstName = rsList("FirstName")
		Set LastName = rsList("LastName")
		Set EMail = rsList("EMail")
		Set MemberEMail = rsList("EMail1")


		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				rsList.Filter = "ID = " & TargetID

				PrintTableData

				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
		set rsPage = Nothing
		set rsList = Nothing
	end if

'They are just cycling through the customers.  No searching.
else
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsPage.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing


	if strVersion = "Free" then
		rsPage.Filter = "Version = 'Free'"
	elseif strVersion = "Gold" then
		rsPage.Filter = "Version = 'Gold'"
	elseif strVersion = "Parent" then
		rsPage.Filter = "Version = 'Parent'"
	elseif strVersion = "Child" then
		rsPage.Filter = "Version = 'Child'"
	elseif strVersion = "Other" then
		rsPage.Filter = "(Version <> 'Free' AND Version <> 'Gold' AND Version <> 'Parent' AND Version <> 'Child')"
	end if

	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="customers.asp">
<%



		Set CustomerID = rsPage("ID")
		Set OwnerID = rsPage("OwnerID")
		Set SignupDate = rsPage("SignupDate")

		Set UseDomain = rsPage("UseDomain")
		Set DomainName = rsPage("DomainName")
		Set SubDirectory = rsPage("SubDirectory")


		Set Title = rsPage("Title")
		Set NickName = rsPage("NickName")
		Set Password = rsPage("Password")
		Set FirstName = rsPage("FirstName")
		Set LastName = rsPage("LastName")
		Set EMail = rsPage("EMail")
		Set MemberEMail = rsPage("EMail1")


		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle
		for j = 1 to rsPage.PageSize
			if not rsPage.EOF then
				PrintTableData
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
	else
%>
			<p>Sorry, but there are no customers at the moment.</p>
<%
	end if
	rsPage.Close
	set rsPage = Nothing
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(CustomerID & OwnerID & MasterID & ParentID & DomainName & SubDirectory & Title & _
	Organization & FirstName & LastName & Street1 & Street2 & City & State & _
	Zip & Country & Phone & BillingType & BillingStreet1 & BillingStreet2 & BillingCity & _
	BillingState & BillingZip & BillingPhone & BillingCountry & CCCompany & CCType & TransID & _
	MerchantClientIDNumber & MerchantBank & MemberFirstName & MemberLastName & HomeStreet & HomeCity & _
	HomeState & HomeZip & HomePhone & Beeper & CellPhone & NickName & Password & _
	EMail & MemberEMail1 & MemberEMail2)
End Function

'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader" align="center">&nbsp;</td>
		<td class="TDHeader" align="center">Address</td>
		<td class="TDHeader" align="center">CustomerID</td>
		<td class="TDHeader" align="center">Signup Date</td>
		<td class="TDHeader" align="center">Subdirectory</td>
		<td class="TDHeader" align="center">Title</td>
		<td class="TDHeader" align="center">Owner Name</td>
	</tr>

<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="button" value="More" onClick="Redirect('customer_view.asp?ID=<%=CustomerID%>')">
		</td>
<%
		if UseDomain = 1 then
			strAddress = DomainName
		else
			strAddress = "http://www.GroupLoop.com/" & SubDirectory
		end if
%>
		<td class="<% PrintTDMain %>"><a href="<%=strAddress%>"><%=strAddress%></a></td>
		<td class="<% PrintTDMain %>"><%=CustomerID%></td>
		<td class="<% PrintTDMain %>"><%=FormatDateTime(SignupDate, 2)%></td>
		<td class="<% PrintTDMain %>"><a href="http://www.GroupLoop.com/<%=SubDirectory%>"><%=SubDirectory%></a></td>
		<td class="<% PrintTDMain %>"><%=Title%></td>
		<td class="<% PrintTDMainSwitch %>"><%=FirstName%>&nbsp;<%=LastName%></td>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->