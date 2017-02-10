<!-- #include file ="adovbs.inc" -->
<%
Set Connect = Server.CreateObject("ADODB.Connection")
strConnection = "PROVIDER=MSDASQL;DRIVER={SQL Server};SERVER=sql1.burlee.com;DATABASE=ourclubpage1;UID=ourclubpage;PWD=hgf554jh;"
Connect.Open strConnection

'This is used for alternating the colors when listing data
Public intBackID
intBackID = 1


'Get the current URL
MyUrl = request.servervariables("SCRIPT_NAME")
MyQuery = request.ServerVariables("QUERY_STRING")
if MyQuery <> "" then MyURl = MyURl & "?" & MyQuery
MyUrl = Right( MyUrl, (Len(MyUrl) - InStrRev(MyUrl,"/") ) )
MyUrl = Left( MyUrl, 50 )

AddHit(MyUrl)

'Response.Write MyUrl




Sub AddHit( strPage )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandType = adCmdStoredProc
		.CommandText = "AddHit"
		.Parameters.Refresh

		if Request("SessionID") <> "" then Session("SessionID") = CInt(Request("SessionID"))

		if Session("SessionID") = "" then Session("SessionID") = 0
		if Session("MemberID") = "" then Session("MemberID") = 0

		.Parameters("@Page") = strPage


		.Parameters("@CustomerID") = CustomerID
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@SessionID") = Session("SessionID")

		.Execute , , adExecuteNoRecords

		if Session("SessionID") = "" then Session("SessionID") = .Parameters("@HitID")

	End With
	Set cmdTemp = Nothing
End Sub

%>
<!-- #include file ="functions.asp" -->
