<%
'-------------------------------------------------------------
'This function gives a quick summary of a customer.  used in lists with items with different customers
'-------------------------------------------------------------
Function GetCustSummary( intCustomerID )
	blKeep = false
	if IsObject(cmd) then blKeep = true

	if not blKeep then Set cmd = Server.CreateObject("ADODB.Command")
	With cmd
		'We already have it loaded...dont do it again
		if .CommandText <> "{ ? = call GetShortCustomerInfo(?, ?, ?, ?, ?, ?) }" then
			.ActiveConnection = Connect
			.CommandText = "GetShortCustomerInfo"
			.CommandType = adCmdStoredProc

			.Parameters.Refresh
		end if

		.Parameters("@CustomerID") = intCustomerID

		.Execute , , adExecuteNoRecords

		strSummary = intCustomerID & " &nbsp; Dir: " & .Parameters("@Subdirectory") & " &nbsp; Version: " & .Parameters("@Version")

	End With
	if not blKeep then Set cmd = Nothing

	GetCustSummary = strSummary
End Function




'-------------------------------------------------------------
'This function updates the price of an invoice with its relative charges
'-------------------------------------------------------------
Function UpdateInvoicePrice( intInvoiceID )
	Set cmdNick = Server.CreateObject("ADODB.Command")
	With cmdNick
		.ActiveConnection = Connect
		.CommandText = "UpdateCalculateCustomerInvoicePrice"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@InvoiceID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Total", adCurrency, adParamOutput )

		.Parameters("@InvoiceID") = intInvoiceID

		.Execute , , adExecuteNoRecords
		currTotal = .Parameters("@Total")
		if not IsNull(currTotal) then currTotal = cDbl(currTotal)
	End With
	Set cmdNick = Nothing

	UpdateInvoicePrice = currTotal
End Function

'-------------------------------------------------------------
'This function returns the nickname of a given Employee
'-------------------------------------------------------------
Function GetEmployeeNickName( intEmployeeID )
	Set cmdNick = Server.CreateObject("ADODB.Command")
	With cmdNick
		.ActiveConnection = Connect
		.CommandText = "GetNickNameEmployee"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Nick", adVarWChar, adParamOutput, 100 )

		.Parameters("@ItemID") = intEmployeeID

		.Execute , , adExecuteNoRecords
		strNick = .Parameters("@Nick")
	End With
	Set cmdNick = Nothing

	GetEmployeeNickName = strNick
End Function


Function EmployeeNickNameTaken( strNickName )
	Set cmdReviews = Server.CreateObject("ADODB.Command")
	With cmdReviews
		.ActiveConnection = Connect
		.CommandText = "EmployeeNickNameTaken"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@NickName", adVarWChar, adParamInput, 100 )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@NickName") = UCase(strNickName)

		.Execute , , adExecuteNoRecords
		blResult = .Parameters("@Exists")
	End With

	Set cmdReviews = Nothing
	EmployeeNickNameTaken = CBool(blResult)
End Function


'-------------------------------------------------------------
'This function logs in a member and sets the Session vars if they are a member
'-------------------------------------------------------------
Sub EmployeeLogin( strPassword, strNickName )
	Set cmdLogin = Server.CreateObject("ADODB.Command")
	With cmdLogin
		.ActiveConnection = Connect
		.CommandText = "ValidEmployee"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@Valid", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@AccessLevel", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@ID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@EmployeeID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@NickName", adVarWChar, adParamInput, 100 )
		.Parameters.Append .CreateParameter ("@Password", adVarWChar, adParamInput, 100 )

		.Parameters("@EmployeeID") = 0
		.Parameters("@NickName") = strNickName
		.Parameters("@Password") = strPassword
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		intAccessLevel = .Parameters("@AccessLevel")
		intEmployeeID = .Parameters("@ID")
	End With
	Set cmdLogin = Nothing

	if blValid then
		Session("NickName") = strNickName
		Session("EmployeeID") = intEmployeeID
		Session("Password") = strPassword
		Session("AccessLevel") = intAccessLevel
		Session("Employee") = "Y"
	end if
End Sub


'-------------------------------------------------------------
'Tells if the person is logged in as a administrator or not
'-------------------------------------------------------------
Function LoggedStaff()
	LoggedStaff = Session("Employee") = "Y"
End Function



'-------------------------------------------------------------
'This function gets an employees first name
'-------------------------------------------------------------
Function GetEmployeeFirstName( intEmployeeID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetEmployeeFirstName"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		if intEmployeeID = 0 then intEmployeeID = Session("EmployeeID")

		.Parameters("@ItemID") = Session("EmployeeID")

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetEmployeeFirstName = strName
End Function


'-------------------------------------------------------------
'This function returns an account name
'-------------------------------------------------------------
Function GetAccountName( intAccountID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetBankAccountName"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@ItemID") = intAccountID

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetAccountName = strName
End Function

'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intHighLight = 0, then it prints a list excluding the intCustMatchID
'if intHighLight = 1, then it prints a list, highlighting the intCustMatchID
'if intChangeSubmit = 1, when they change the option it submits the form
'strname = name of pulldown menu
'strTable = table name from sql
'strPrintBlank = if this is not "", print the label - None, All, whatever
'-------------------------------------------------------------
Sub PrintCustomerPullDown( intCustMatchID, intHighLightID, intChangeSubmit, strPrintBlank, strName )

	if intCustMatchID <> "" then
		intCustMatchID = CInt(intCustMatchID)
	else
		intCustMatchID = 0
	end if


	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 150

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsPulldown.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing


	if rsPulldown.EOF then
		Set rsPulldown = Nothing
		Exit Sub
	end if

	Set CustID = rsPulldown("ID")
	Set Version = rsPulldown("Version")

	Set SubDirectory = rsPulldown("SubDirectory")

	//They don't have to pass the name, we will provide it
	if strName = "" then strName = "CustomerID"

	%><select name="<%=strName%>" size="1" <%

	if intChangeSubmit = 1 then
		%>onChange="this.form.submit();"<%
	end if

	Response.Write ">"
	if strPrintBlank <> "" then Response.Write "<option value = ''>" & strPrintBlank & "</option>" & vbCrlf

	do until rsPulldown.EOF
		'Highlight the current category
		if intHighLightID = 1 and CustID = intCustMatchID then
			Response.Write "<option value = '" & CustID & "' SELECTED>ID #" & CustID & " Dir: " & SubDirectory & " Version: " & Version & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & CustID & "'>ID #" & CustID & " Dir: " & SubDirectory & " Version: " & Version & "</option>" & vbCrlf
		end if

		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

End Sub




'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intChangeSubmit = 1, when they change the option it submits the form
'strname = name of pulldown menu
'-------------------------------------------------------------
Sub PrintAccountsPullDown( intHighLightID, strName )

	Query = "SELECT ID, Description, AccountNumber FROM BankAccounts"
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 5
	rsPulldown.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsPulldown.EOF then Exit Sub

	Set AcctID = rsPulldown("ID")
	Set Description = rsPulldown("Description")
	Set AccountNumber = rsPulldown("AccountNumber")

	//They don't have to pass the name, we will provide it
	if strName = "" then strName = "AccountID"

	%><select name="<%=strName%>" size="1" <%

	Response.Write ">"

	do until rsPulldown.EOF
		'Highlight the current category
		if AcctID = intHighLightID then
			Response.Write "<option value = '" & AcctID & "' SELECTED>" & Description & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & AcctID & "'>" & Description & "</option>" & vbCrlf
		end if

		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

End Sub



'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the bank statements
'if intChangeSubmit = 1, when they change the option it submits the form
'strname = name of pulldown menu
'-------------------------------------------------------------
Sub PrintStatementsPullDown( intHighLightID, strName )

	Query = "SELECT ID, AccountID, DateStarted, DateEnded FROM BankStatements"
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 5
	rsPulldown.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsPulldown.EOF then Exit Sub


	Set SID = rsPulldown("ID")
	Set AccountID = rsPulldown("AccountID")
	Set DateStarted = rsPulldown("DateStarted")
	Set DateEnded = rsPulldown("DateEnded")

	//They don't have to pass the name, we will provide it
	if strName = "" then strName = "AccountID"

	%><select name="<%=strName%>" size="1" <%

	Response.Write ">"

	strSelected = ""
	if intHighLightID = 0 then strSelected = " SELECTED"

	Response.Write "<option value = ''" & strSelected & ">None</option>" & vbCrlf

	do until rsPulldown.EOF
		'Highlight the current category
		if SID = intHighLightID then
			Response.Write "<option value = '" & SID & "' SELECTED>" & GetAccountName( AccountID ) & FormatDateTime(DateStarted, 2) & " - " & FormatDateTime(DateEnded, 2) & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & SID & "'>" & GetAccountName( AccountID ) & FormatDateTime(DateStarted, 2) & " - " & FormatDateTime(DateEnded, 2) & "</option>" & vbCrlf
		end if

		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

End Sub
%>	