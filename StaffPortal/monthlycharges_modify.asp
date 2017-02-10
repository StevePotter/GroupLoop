<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Monthly Charges</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if


if not LoggedStaff() then Redirect("login.asp?Source=monthlycharges_modify.asp&ID=" & Request("ID"))

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Date, Total, Description, CustomerID FROM CustomerMonthlyCharges WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	intCustomerID = rsUpdate("CustomerID")

	rsUpdate("Date") = AssembleDate( "Date" )
	rsUpdate("Total") = cDbl(Request("Total"))
	rsUpdate("Description") = Format( Request("Description") )


	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The monthly charge has been edited.  
	<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.</a><br>
	<a href="customers.asp">Browse the list of customers.</a>
	</p>

<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, CustomerID FROM CustomerMonthlyCharges WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	intCustomerID = rsUpdate("CustomerID")
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The monthly charge has been deleted. 
	<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.</a><br>
	<a href="customers.asp">Browse the list of customers.</a>
	</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Total, Description FROM CustomerMonthlyCharges WHERE ID = " & intID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

'------------------------End Code-----------------------------
%>
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


		function submit_page(form) {
			//Error message variable
			var strError = "";
			form.Total.value = ConvertDollar(form.Total.value)

			if (form.Total.value == "" || form.Total.value == "0.00" || form.Total.value == "0" )
				strError += "          You forgot the total. \n";

			if (form.Description.value == "" )
				strError += "          You forgot the description. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can add this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>

	* indicates required information<br>
	<form method="post" action="monthlycharges_modify.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
		<% DatePulldown "Date", rsEdit("Date"), 1 %>
     	</td>
    </tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">* Charge Total</td>
			<td class="<% PrintTDMain %>"> 
				<input type="text" name="Total" value="<%=FormatCurrency(rsEdit("Total"))%>" size="5">
			</td>
   		</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Description</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="Description" cols="55" rows="5" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Description") )%></textarea>
    	</td>
	</tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
	</tr>

  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

else
	Query = "SELECT ID, Date, Total, CustomerID, Description FROM CustomerMonthlyCharges ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Total = rsPage("Total")
		Set rsCustomerID = rsPage("CustomerID")
		Set Description = rsPage("Description")

		Set cmd = Server.CreateObject("ADODB.Command")	'used for the GetCustSummary.  this way we create/destroy object once
'-----------------------End Code----------------------------
%>
		<a href="monthlycharge_add.asp">Add New Monthly Charge</a><br>
		<form METHOD="POST" ACTION="monthlycharges_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Customer</td>
			<td class="TDHeader">Total</td>
			<td class="TDHeader">Description</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="monthlycharges_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>"><%=GetCustSummary( rsCustomerID )%></td>
						<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
						<td class="<% PrintTDMain %>"><%=Description%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this charge, there is no way to get it back.  Are you sure?', 'monthlycharges_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close

		Set cmd = Nothing
	else
'------------------------End Code-----------------------------
%>
		<p>You have to create monthly charges for you can modify them, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if

%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->