<!-- #include file="dsn.asp" -->

<!-- #include file="header_blank.asp" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("InvoiceID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the InvoiceID."))
intInvoiceID = CInt(Request("InvoiceID"))

if not LoggedStaff() then Redirect("login.asp?Source=invoice_print.asp&ID=" & Request("InvoiceID"))

Query = "SELECT * FROM CustomerInvoices WHERE ID = " & intInvoiceID
Set rsInvoice = Server.CreateObject("ADODB.Recordset")
rsInvoice.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

if rsInvoice.EOF then
	Set rsInvoice = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
end if



Query = "SELECT Date, Total, DateDeposited FROM BankDeposits WHERE InvoiceID = " & intInvoiceID
Set rsDeposits = Server.CreateObject("ADODB.Recordset")
rsDeposits.CacheSize = 100
rsDeposits.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

blDeposits = not rsDeposits.EOF   'If there are records, we have deposits, if there aren't, we don't

dblDue = rsInvoice("Total")

if blDeposits then
	do until rsDeposits.EOF
		dblDue = dblDue - rsDeposits("Total")
		rsDeposits.MoveNext
	loop

	rsDeposits.MoveFirst
end if

intCustomerID = rsInvoice("CustomerID")

Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.CacheSize = 100

Set cmdTemp = Server.CreateObject("ADODB.Command")
cmdTemp.ActiveConnection = Connect
cmdTemp.CommandText = "GetSiteInfoRecordSet"
cmdTemp.CommandType = adCmdStoredProc

rsPage.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

Set cmdTemp = Nothing

rsPage.Filter = "ID = " & intCustomerID


if rsPage("UseDomain") = 1 then
	strAddress = rsPage("DomainName")
else
	strAddress = "http://www.GroupLoop.com/" & rsPage("SubDirectory")
end if

if IsNull(rsInvoice("DateSent")) then
	dateSent = FormatDateTime( rsInvoice("Date") , 2 )
else
	dateSent = FormatDateTime( rsInvoice("DateSent") , 2 )
end if
%>


<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0><br>
GroupLoop.com<br>
P.O. Box 5271<br>
Somerville, NJ 08876-3430<br>
accounts@grouploop.com<br>
<hr noshade align="center" width="300"><br>
<div align="center"><b><font face="Arial, Helvetica, sans-serif" size="+4">Invoice</font></b></div>
<table  border="1" cellspacing="0" cellpadding="4" bordercolor="#000000" align="right">
	<tr> 
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
		Date
		</font></b>
	</td>
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
		Invoice #
		</font></b>
	</td>
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
		Customer #
		</font></b>
	</td>
	</tr>
	<tr> 
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
		<%=dateSent%>
		</font></b>
	</td>
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
	<%=intInvoiceID%>
		</font></b>
	</td>
	<td align="center"><b>
		<font face="Arial, Helvetica, sans-serif">
	<%=intCustomerID%>
		</font></b>
	</td>
	</tr>
</table>
</p>

<table width="300" border="1" cellspacing="0" cellpadding="4" bordercolor="#000000">
<tr> 
<td>
<b>Bill To:</b><br>

<%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%><br>
<%
if rsPage("Organization") <> "" then
%>
<%=rsPage("Organization")%><br>
<%
end if%>
<%=rsPage("Street1")%><br>
<%
if rsPage("Street2") <> "" then
%>
	<%=rsPage("Street2")%><br>
<%
end if
%>
<%=rsPage("City")%>,&nbsp;<%=rsPage("State")%>&nbsp;<%=rsPage("Zip")%>&nbsp;<%=rsPage("Country")%><br>
<%
if rsPage("Phone") <> "" then
%>
	<%=rsPage("Phone")%><br>
<%
end if
%>

</td>
</tr>
</table>

<p>
<b><font face="Arial, Helvetica, sans-serif" size="+1">
<%
if rsInvoice("Description") <> "" then
%>
Invoice Description: <%=rsInvoice("Description")%><br>
<% end if %>
Total Due: <%=FormatCurrency(dblDue)%></font>
</b>
</p>
<%
	Query = "SELECT ID, Date, Hours, Description, StaffNote, CustomerNote, DateStarted, DateEnded, Total FROM CustomerInvoiceCharges WHERE InvoiceID = " & intInvoiceID & " ORDER BY ID"
	Set rsCharges = Server.CreateObject("ADODB.Recordset")
	rsCharges.CacheSize = PageSize
	rsCharges.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsCharges.EOF then

		Set ID = rsCharges("ID")
		Set ItemDate = rsCharges("Date")
		Set Hours = rsCharges("Hours")
		Set Description = rsCharges("Description")
		Set StaffNote = rsCharges("StaffNote")
		Set CustomerNote = rsCharges("CustomerNote")
		Set DateStarted = rsCharges("DateStarted")
		Set DateEnded = rsCharges("DateEnded")
		Set Total = rsCharges("Total")

%>

		<table width="100%" border="1" cellspacing="0" cellpadding="4" bordercolor="#000000">
		<tr> 
		<td width="80%" align="center">
			<b><font face="Arial, Helvetica, sans-serif">
			Description
			</font></b>
		</td>
		<td width="20%" align="center">
			<b><font face="Arial, Helvetica, sans-serif">
			Amount
			</font></b>			
		</td>
		</tr>
<%
		do until rsCharges.EOF
			strDescription = FormatDateTime(ItemDate, 2) & ":&nbsp; " & Description

			if Hours > 0 then strDescription = strDescription & "<br>" & Hours & " hour" & PrintPlural(Hours, "", "s") & " spent."

			if CustomerNote <> "" then strDescription = strDescription & "<br>" & "Note:&nbsp; " & CustomerNote

'------------------------End Code-----------------------------
%>
			<tr> 
			<td width="80%" valign="middle">
			<font face="Arial, Helvetica, sans-serif">
				<%=strDescription%>
			</font>
			</td>
			<td width="20%" valign="middle">
			<font face="Arial, Helvetica, sans-serif">
				<%=FormatCurrency(Total)%>
			</font>		
			</td>
			</tr>
<%
'-----------------------Begin Code----------------------------
			rsCharges.MoveNext
		loop

			if blDeposits then
%>
			<tr> 
			<td width="80%" align="right"><b>
			<font face="Arial, Helvetica, sans-serif">
				Sub-Total &nbsp;
			</b></font>
			</td>
			<td width="20%"><b>
			<font face="Arial, Helvetica, sans-serif">
				<%=FormatCurrency(rsInvoice("Total"))%>
			</b></font>		
			</td>
			</tr>
<%
				do until rsDeposits.EOF
%>
					<tr> 
					<td width="80%" align="right"><b>
					<font face="Arial, Helvetica, sans-serif">
						Payment on <%=FormatDateTime(rsDeposits("Date"), 2)%>  &nbsp;
					</b></font>
					</td>
					<td width="20%"><b>
					<font face="Arial, Helvetica, sans-serif">
						-<%=FormatCurrency(rsDeposits("Total"))%>
					</b></font>		
					</td>
					</tr>


<%
					rsDeposits.MoveNext
				loop
%>
					<tr> 
					<td width="80%" align="right"><b>
					<font face="Arial, Helvetica, sans-serif">
						TOTAL DUE  &nbsp;
					</b></font>
					</td>
					<td width="20%"><b>
					<font face="Arial, Helvetica, sans-serif">
						<%=FormatCurrency(dblDue)%>
					</b></font>		
					</td>
					</tr>


<%

			else
%>
			<tr> 
			<td width="80%" align="right"><b>
			<font face="Arial, Helvetica, sans-serif">
				TOTAL &nbsp;
			</b></font>
			</td>
			<td width="20%"><b>
			<font face="Arial, Helvetica, sans-serif">
				<%=FormatCurrency(rsInvoice("Total"))%>
			</b></font>		
			</td>
			</tr>
<%
			end if
		Response.Write("</table>")
	end if
%>
<p>
<font face="Arial, Helvetica, sans-serif" size="+1">
<u><b>Make payable to:</b></u><blockquote>
</font>
<font size="+1">
<b>
GroupLoop.com<br>
P.O. Box 5271<br>
Somerville, NJ 08876-3430<br>
</font>
</blockquote>
</b>
</p>

<%
if rsInvoice("CustomerNote") <> "" then
%>
<font face="Arial, Helvetica, sans-serif" size="+1">
<u><b>Note:</b></u>
</font>	
<blockquote><%=rsInvoice("CustomerNote")%></blockquote>
<%
end if 
'-----------------------Begin Code----------------------------
	rsPage.Close
	Set rsPage = Nothing
	rsCharges.Close
	Set rsCharges = Nothing
	rsInvoice.Close
	set rsInvoice = Nothing

%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->