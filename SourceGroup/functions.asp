<!-- #include file="..\sourcecommon\functions.asp" -->

<%
'This file must have the dsn.asp file and CustomerID included before it!
Function GetNumItems( strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumItems"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@Table", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )
		.Parameters("@Table") = strTable
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumItems = intCount
End Function


Function GetCommittee( intCommitteeID )
	if intCommitteeID = "" then intCommitteeID = 0
	intCommitteeID = CInt(intCommitteeID)
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCommittee"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 200 )
		.Parameters("@ItemID") = intCommitteeID
		.Execute , , adExecuteNoRecords
		strTempName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetCommittee = strTempName
End Function

'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intHighLight = 0, then it prints a list
'if intHighLight > 1, then it prints a list, highlighting the intCategoryID
'-------------------------------------------------------------
Sub PrintCommitteePullDown( intHighLightID )

	intHighLightID = CInt(intHighLightID)

	Query = "SELECT Name, ID FROM Committees WHERE CustomerID = " & CustomerID & " ORDER BY Name"

	Set rsTempCats = Server.CreateObject("ADODB.Recordset")
	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsTempCats.EOF then
		Set rsTempCats = Nothing
		Exit Sub
	end if

	Set Name = rsTempCats("Name")
	Set ID = rsTempCats("ID")

	'Make the size 3 if there are many committees
	if rsTempCats.RecordCount <= 30 then
		%><select name="CommitteeID" size="1"><%
	else
		%><select name="CommitteeID" size="3"><%
	end if

	Response.Write "<option value='0'>None</option>"

	do until rsTempCats.EOF
		if ID = intHighLightID then
			Response.Write "<option value = '" & ID & "' SELECTED>" & Name & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & ID & "'>" & Name & "</option>" & vbCrlf
		end if

		rsTempCats.MoveNext
	loop
	rsTempCats.Close
	set rsTempCats = Nothing

	Response.Write("</select>")
End Sub


'-------------------------------------------------------------
'This function says if they are a member of more than one site
'-------------------------------------------------------------
Function MultiSiteMember()
	blMultiSites = ParentSiteExists() or ChildSiteExists()

	if not blMultiSites then
		MultiSiteMember = false
	end if

	intMemberID = CInt(Session("MemberID"))
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "IsMultiSiteMember"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Valid", adInteger, adParamOutput )
		.Parameters("@MemberID") = intMemberID
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
	End With
	Set cmdTemp = Nothing

	MultiSiteMember = blValid
End Function


'-------------------------------------------------------------
'This sub gets the recordset of the different sites this person is a member of (linked by commonID)
'-------------------------------------------------------------
Sub GetMemberSitesRecordset( rsSites )
	rsSites.CacheSize = 20

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetMemberSites"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )

		.Parameters("@MemberID") = CInt(Session("MemberID"))

		rsSites.Open cmdTemp, , adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		Response.write sitecount

	End With
	Set cmdTemp = Nothing

End Sub

'-------------------------------------------------------------
'This sub prints a pulldown of the different sites this person is a member of (linked by commonID)
'-------------------------------------------------------------
Sub PrintMemberSites()
	Set rsSites = Server.CreateObject("ADODB.Recordset")

	GetMemberSitesRecordset rsSites

	if rsSites.EOF then
		Set rsSites = Nothing
		Response.Write "You aren't a member of any sites"
		Exit Sub
	end if

	%><select name="SiteCustID" size="3" multiple><%

	Response.Write "<option value='All'>All Sites</option>"

	if not rsSites.EOF then
		Set SiteCustID = rsSites("CustomerID")
		Set SiteTitle = rsSites("Title")
	end if

	do until rsSites.EOF
		if SiteCustID = CustomerID then
			Response.Write "<option value = '" & SiteCustID & "' SELECTED>" & SiteTitle & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & SiteCustID & "'>" & SiteTitle & "</option>" & vbCrlf
		end if

		rsSites.MoveNext
	loop
	rsSites.Close
	set rsSites = Nothing

	Response.Write("</select>")
End Sub

Function ParentSiteExists()

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ParentSiteExists"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	ParentSiteExists = blValid
End Function

Function ChildSiteExists()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ChildSiteExists"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	ChildSiteExists = blValid
End Function

Function GetSiteTitle( TempCustID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetSiteTitle"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Title", adVarWChar, adParamOutput, 200 )
		.Parameters("@CustomerID") = TempCustID
		.Execute , , adExecuteNoRecords
		strTitle = .Parameters("@Title")
	End With
	Set cmdTemp = Nothing

	GetSiteTitle = strTitle
End Function
%>
