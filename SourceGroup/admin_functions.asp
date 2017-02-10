<%
'-------------------------------------------------------------
'See if the administrator wants the section to be displayed
'-------------------------------------------------------------
Function CheckLook( strSection )
	Query = "SELECT " & strSection & " FROM Look WHERE CustomerID = " & CustomerID
	Set rsLookTest = Server.CreateObject("ADODB.Recordset")
	rsLookTest.Open Query, Connect, adOpenStatic, adLockReadOnly
	intCheck = rsLookTest( strSection )
	rsLookTest.Close
	set rsLookTest = Nothing
	CheckLook = ( intCheck = 1 )
End Function

'-------------------------------------------------------------
'This function gets the color from a page that someone chose a color from (using the pulldown, radio, or text box)
'-------------------------------------------------------------
Function GetColor( strRequest )
	if Request(strRequest & "Cust") <> "" and Request(strRequest) = "Custom" then
		GetColor = Request(strRequest & "Cust")
	else
		GetColor = Request(strRequest)
	end if
End Function
%>