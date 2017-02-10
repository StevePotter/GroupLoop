<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add Monthly Charge</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=monthlycharge_add.asp&CustomerID=" & Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Add" then
	if Request("CustomerID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the CustomerID."))
	intCustomerID = CInt(Request("CustomerID"))


	Set cmdReviews = Server.CreateObject("ADODB.Command")
	With cmdReviews
		.ActiveConnection = Connect
		.CommandText = "AddCustomerMonthlyCharge"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@Total") = cDbl(Request("Total"))
		.Parameters("@Description") = Request("Description")

		.Execute , , adExecuteNoRecords
	End With

	Set cmdReviews = Nothing
%>
<p>The monthly charge has been added.  
<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.</a><br>
<a href="customers.asp">Browse the list of customers.</a>
</p>

<%
else
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
	<form method="post" action="monthlycharge_add.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
<%
	if Request("CustomerID") = "" then
%>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Customer</td>
     	<td class="<% PrintTDMain %>"> 
    		<% PrintCustomerPullDown 0, 1, 0, "", "" %>
    	</td>
	</tr>
<%
	else
		intCustomerID = CInt(Request("CustomerID"))
%>
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
<%
	end if
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">* Charge Total</td>
			<td class="<% PrintTDMain %>"> 
				<input type="text" name="Total" value="$" size="5">
			</td>
   		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Description</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Description" cols="55" rows="5" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
<%
end if
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->