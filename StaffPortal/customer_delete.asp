<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Delete Customer</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'-----------------------Begin Code----------------------------
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if not LoggedStaff() then Redirect("login.asp?Source=customer_delete.asp&ID=" & intID)

if Request("Submit") = "Yes" then
	intID = CInt(Request("ID"))

	Response.Write DeleteCustomer(intID)

'------------------------End Code-----------------------------
%>
	<p>
	The customer has been deleted. &nbsp;<a href="customers.asp">Click here to browse the customer list.</a><br>
	</p>
<%
'-----------------------Begin Code----------------------------
elseif Request("Submit") = "No" then
%>
	<p>
	You have chosen not to remove the customer. &nbsp;<a href="customer_view.asp?ID=<%=intID%>">Click here</a> to view the customer's information.</a><br>
	<a href="customers.asp">Click here</a> to browse the customer list.</a>
	</p>
<%
else
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsPage.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	rsPage.Filter = "ID = " & intID

	if rsPage.EOF then Redirect("error.asp?Message=" & Server.URLEncode("The customer has been deleted from the database."))


	if rsPage("UseDomain") = 1 then
		strAddress = rsPage("DomainName")
	else
		strAddress = "http://www.GroupLoop.com/" & rsPage("SubDirectory")
	end if

'------------------------End Code-----------------------------
%>
	<form METHOD="POST" ACTION="customer_delete.asp">
	<input type="hidden" name="ID" value="<%=intID%>">

	<p>
	<i>You have chosen to remove this customer:</i><br>
	&nbsp;&nbsp;&nbsp;Customer ID: <%=intID%><br>
	&nbsp;&nbsp;&nbsp;Site Address: <a href="<%=strAddress%>"><%=strAddress%></a><br>
	&nbsp;&nbsp;&nbsp;Date Created: <%=FormatDateTime(rsPage("SignupDate"), 2)%><br>
	&nbsp;&nbsp;&nbsp;Owner Name: <%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%><br>
	&nbsp;&nbsp;&nbsp;Site Title: <%=rsPage("Title")%><br>
	&nbsp;&nbsp;&nbsp;Contact E-Mail: <a href="mailto:<%=rsPage("EMail")%>"><%=rsPage("EMail")%></a>
	</p>

	<p><b>Are you sure?</b><br>
	<input type="submit" name="submit" value="Yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="submit" name="submit" value="No">

	</form>
<%
	rsPage.Close
	Set rsPage = Nothing
end if


	Function DeleteCustomer( intCustomerID )

		strReturn = ""

		Set Command = Server.CreateObject("ADODB.Command")
		With Command
			'Get the subdirectory
			.ActiveConnection = Connect
			.CommandText = "GetCustomerInfo"
			.CommandType = adCmdStoredProc
			.Parameters.Refresh
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			strSubDir = .Parameters("@SubDirectory")

			.CommandText = "DeleteCustomer"
			.Parameters.Refresh
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			strReturn = strReturn &  "1. Removing from database<br>"
		End With
		Set Command = Nothing

		if strSubDir = "" then Redirect("error.asp?Message=" & Server.URLEncode("The customer has been deleted from the database, but the directory didn't exist."))
		strFolder = Server.MapPath("../" & strSubDir)

		if strFolder = "E:\Webs\Websites\ourclubpage.com" then
			Redirect "error.asp"
		end if

		Set FSys = CreateObject("Scripting.FileSystemObject")
		if FSys.FolderExists(strFolder) then
			strReturn = strReturn &  "2. Removing folder: " & strFolder & "<br>"
			FSys.DeleteFolder strFolder
		else
			strReturn = strReturn &  "2. FOLDER NOT FOUND - " & strFolder & "<br>"
		end if
		Set FSys = Nothing

		DeleteCustomer = strReturn
	End Function
'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->