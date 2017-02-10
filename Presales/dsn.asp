<!-- #include file ="..\sourcegroup\adovbs.inc" -->

<%
Set Connect = Server.CreateObject("ADODB.Connection")
strConnection = "PROVIDER=MSDASQL;DRIVER={SQL Server}; SERVER=sql1.burlee.com;DATABASE=ourclubpage1;UID=ourclubpage;PWD=hgf554jh;"
Connect.Open strConnection

Sub AddHit( strPage )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddHit"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh

		if Request("SessionID") <> "" then Session("SessionID") = CInt(Request("SessionID"))

		if Session("SessionID") = "" then Session("SessionID") = 0
		if Session("MemberID") = "" then Session("MemberID") = 0

		.Parameters("@Page") = strPage


		.Parameters("@CustomerID") = -1
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@SessionID") = Session("SessionID")

		.Execute , , adExecuteNoRecords

		if Session("SessionID") = "" then Session("SessionID") = .Parameters("@HitID")

	End With
	Set cmdTemp = Nothing
End Sub
%>
