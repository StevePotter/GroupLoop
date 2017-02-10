<!-- #include file="..\sourcegroup\functions.asp" -->

<%
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
		Session("EmployeeID") = intEmployeeID
		Session("Password") = strPassword
		Session("AccessLevel") = intAccessLevel
		Session("Employee") = "Y"
	end if
End Sub


'-------------------------------------------------------------
'Tells if the person is logged in as a administrator or not
'-------------------------------------------------------------
Function LoggedEmployee( )
	LoggedEmployee = Session("Employee") = "Y"
End Function



'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function GetEmployeeFirstName( intEmployeeID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetEmployeeFirstName"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 100 )

		if intEmployeeID = 0 then intEmployeeID = Session("EmployeeID")

		.Parameters("@ItemID") = Session("EmployeeID")

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetEmployeeFirstName = strName
End Function



%>
