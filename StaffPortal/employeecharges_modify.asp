<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Paycheck Charges</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if


if not LoggedStaff() then Redirect("login.asp?Source=employeecharges_modify.asp&ID=" & Request("ID"))

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" or Request("PaycheckID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))
	intPaycheckID = CInt(Request("PaycheckID"))

	Set cmdReviews = Server.CreateObject("ADODB.Command")
	With cmdReviews
		.ActiveConnection = Connect
		.CommandText = "UpdateCustomerPaycheckCharge"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		.Parameters("@ChargeID") = intID

		.Parameters("@Total") = cDbl(Request("Total"))
		if Request("Hours") <> "" then .Parameters("@Hours") = cDbl(Request("Hours"))
		.Parameters("@Description") = Format(Request("Description"))
		.Parameters("@CustomerNote") = Format(Request("CustomerNote"))
		.Parameters("@StaffNote") = Format(Request("StaffNote"))
		.Parameters("@PaycheckID") = intPaycheckID

		'Custom time included
		DateStarted = AssembleDate("DateStarted")

		DateEnded = AssembleDate("DateEnded")

		.Parameters("@DateStarted") = DateStarted
		.Parameters("@DateEnded") = DateEnded

		.Execute , , adExecuteNoRecords

		intCustomerID = .Parameters("@CustomerID")

	End With

	Set cmdReviews = Nothing

	'update the paycheck price
	currTemp = UpdatePaycheckPrice( intPaycheckID )

	'if we swtiched to a new paycheck, update the old price
	intOldPaycheck = CInt(Request("OldPaycheckID"))
	if intOldPaycheck <> intPaycheckID then currTemp = UpdatePaycheckPrice( intOldPaycheck )



'------------------------End Code-----------------------------
%>
<p>The paycheck charge has been edited.<br>
<a href="paychecks_modify.asp?Submit=Edit&ID=<%=intPaycheckID%>">View the paycheck the charge belongs to.<br>
<a href="customer_view.asp?ID=<%=intCustomerID%>">View this customer.<br>
</p>

<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, PaycheckID, CustomerID FROM EmployeePaycheckChrages WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	intPaycheckID = rsUpdate("PaycheckID")
	CustomerID = rsUpdate("CustomerID")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	currTemp = UpdatePaycheckPrice( intPaycheckID )
'------------------------End Code-----------------------------
%>
	<p>The paycheck charge has been deleted.  <a href="paychecks_modify.asp?Submit=Edit&ID=<%=intPaycheckID%>">View the paycheck</a><br>
	<a href="customer_view.asp?ID=<%=CustomerID%>">View the customer's details.</a><br>
	<a href="customers.asp">Browse the list of customers.</a>
	</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM EmployeePaycheckChrages WHERE ID = " & intID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	intPaycheckID = rsEdit("PaycheckID")

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

		function AssembleDate(form, Field ){
			var intDay = form.elements[Field + 'Day'].value;
			var intMonth = form.elements[Field + 'Month'].value;
			var intYear = form.elements[Field + 'Year'].value;

			//convert to numerics
			intDay -= 0;
			intMonth -= 0;
			intMonth -= 1;	//must decrement month
			intYear -= 0;

			var strTime = form.elements[Field + 'Time'].value;
			//the first half will have the hh:mm:ss, second will have am or pm
			var dayHalf = strTime.split(' ');
			var strFullTime = dayHalf[0];
			var AMPM = dayHalf[1];
			var strTime = strFullTime.split(':');

			var intHour = strTime[0];
			intHour -= 0;

			if ( AMPM == 'AM' ){
				//12 am is really 0
				if ( intHour == 12 ){
					intHour = 0;
				}
			}
			else{
				if ( intHour < 12 )	//12 pm is really 12, so leave it 
					intHour += 12;
			}

			intHour -= 0;
			var intMin = strTime[1]
			intMin -= 0;

			var intSec = strTime[2]
			intSec -= 0;

			var date = new Date();

			date.setDate(intDay);
			date.setMonth(intMonth);
			date.setYear(intYear);
			date.setHours(intHour);
			date.setMinutes(intMin);
			date.setSeconds(intSec);


			return date;
		}

		function PutDate(form, Field){
			var date = new Date();
			var d  = date.getDate();
			var day = (d < 10) ? '0' + d : d;
			var m = date.getMonth() + 1;
			var month = (m < 10) ? '0' + m : m;
			var yy = date.getYear();
			var year = (yy < 1000) ? yy + 1900 : yy;



			myhours = date.getHours();
			if (myhours >= 12) {
			myhours = (myhours == 12) ? 12 : myhours - 12; mm = " PM";
			}
			else {
			myhours = (myhours == 0) ? 12 : myhours; mm = " AM";
			}
			myminutes = date.getMinutes();
			if (myminutes < 10){
			myminutes = ":0" + myminutes;
			}
			else {
			myminutes = ":" + myminutes;
			};
			mysecs = date.getSeconds();
			if (mysecs < 10){
			mysecs = ":0" + mysecs;
			}
			else {
			mysecs = ":" + mysecs;
			};

			form.elements[Field + 'Month'].value = m;
			form.elements[Field + 'Day'].value = d;
			form.elements[Field + 'Year'].value = year;


			if ( form.elements[Field + 'Time'] )
				form.elements[Field + 'Time'].value = myhours+myminutes+mysecs+mm;

			return;
		}

		function StartWork(form){
			
			updateClocks();

			return;

		}

		function StopWork(){
			//Stop the running clock
			clearTimeout(timeoutID);

			CalculateCost();

			return;
		}

		function updateClocks() {
			now = new Date();
			PutDate( document.MyForm, 'DateEnded');

			CalculateHours(document.MyForm, 'DateStarted', 'DateEnded', 'Hours');
			CalculateCost();

			timeoutID = setTimeout('updateClocks()',500);
			return;
		}

		//Clock ID
		var timeoutID = 0;


		function CalculateHours(form, Field1, Field2, DisplayField) {
			if (form.elements[Field1+'Day'].value == '' || form.elements[Field2+'Day'].value == '')
				return;

			var earlierdate = AssembleDate(form, Field1 );
			var laterdate = AssembleDate(form, Field2 );

		    var difference = laterdate.getTime() - earlierdate.getTime();

			var Secs = Math.floor(difference/1000);

			//900 secs every quarter hour.  get the whole number of quarter hours (round up)
			var unroundedquarterHour = (Secs/900);			
			var roundedquarterHour = Math.floor(unroundedquarterHour);
			//if there is a part of a quarter hour left, increment the number
			if ( unroundedquarterHour > roundedquarterHour )
				roundedquarterHour ++;

			varHours = (roundedquarterHour/4);

			form.elements[DisplayField].value = varHours;
		}

		function CalculateCost() {
			var form = document.MyForm;
			CalculateHours(form, 'DateStarted', 'DateEnded', 'Hours');

			//if there is no hours worked, we can't put a cost
			if (form.elements['Hours'].value == '' || form.elements['Hours'].value == '0'){
				form.elements['Total'].value = '$0.00';
				return;
			}

			//Get the current rate
			form.elements['Rate'].value = ConvertDollar(form.elements['Rate'].value)
			var intRate = form.elements['Rate'].value;
			if ( intRate == '' ){
				form.elements['Rate'].value = '60';
				intRate = form.elements['Rate'].value;
			}
			intRate -= 0;

			var intHours = form.elements['Hours'].value;
			intHours -= 0;

			var Total = (intHours * intRate);

			
			form.elements['Total'].value = outputMoney(Total);
		}



		function outputMoney(number) {
			return '$' + outputDollars(Math.floor(number-0) + '') + outputCents(number - 0);
		}

		function outputDollars(number) {
			if (number.length <= 3)
				return (number == '' ? '0' : number);
			else {
				var mod = number.length%3;
				var output = (mod == 0 ? '' : (number.substring(0,mod)));
				for (i=0 ; i < Math.floor(number.length/3) ; i++) {
					if ((mod ==0) && (i ==0))
						output+= number.substring(mod+3*i,mod+3*i+3);
					else
						output+= ',' + number.substring(mod+3*i,mod+3*i+3);
				}
				return (output);
			}
		}

		function outputCents(amount) {
			amount = Math.round( ( (amount) - Math.floor(amount) ) *100);
			return (amount < 10 ? '.0' + amount : '.' + amount);
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
	<form method="post" action="employeecharges_modify.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<input type="hidden" name="OldPaycheckID" value="<%=intPaycheckID%>">
	<input type="button" value="Calculate Cost" onClick="CalculateCost()">&nbsp;&nbsp;&nbsp;
	<input type="button" value="Start" onClick="StartWork()">&nbsp;&nbsp;&nbsp;
	<input type="button" value="Stop" onClick="StopWork()">

<%
	if rsEdit("Total") > 0 and rsEdit("Hours") > 0 then
		Rate = FormatCurrency( rsEdit("Total")/rsEdit("Hours") )
	else
		Rate = "$"
	end if
%>

	<%PrintTableHeader 0%>


	<tr> 
		<td class="<% PrintTDMain %>" align="right">* Paycheck</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintPaycheckPullDown intPaycheckID %>
		</td>
	</tr>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">* Charge Total</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="Total" value="<%=FormatCurrency(rsEdit("Total"))%>" size="5">&nbsp;&nbsp;&nbsp;@ <input type="text" name="Rate" value="<%=Rate%>" size="5"> per hour.
		</td>
	</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Created</td>
      	<td class="<% PrintTDMain %>"> 
		<% DatePulldown "Date", rsEdit("Date"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'Date')">
		</td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Work Started</td>
      	<td class="<% PrintTDMain %>"> 
		<% DatePulldown "DateStarted", rsEdit("DateStarted"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateStarted')">
		</td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Work Ended</td>
      	<td class="<% PrintTDMain %>"> 
		<% DatePulldown "DateEnded", rsEdit("DateEnded"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateEnded')">
		</td>
    </tr>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Hours - 15 min intervals</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="Hours" value="<%=rsEdit("Hours")%>" size="5">
		</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Description</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Description") )%></textarea>
    	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Any Special Note/Reminder</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="CustomerNote" cols="55" rows="2" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Note") )%></textarea>
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
	Query = "SELECT ID, Date, Name FROM EmployeePaycheckChrages ORDER BY CustomerID, ID"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Name = rsPage("Name")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="employeecharges_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Topic</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="employeecharges_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="forum.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Name%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this topic, there is no way to get it or its messages back.  Are you sure?', 'employeecharges_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
	else
'------------------------End Code-----------------------------
%>
		<p>You have to create paycheck charges for you can modify them, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if



Sub PrintPaycheckPullDown( intHighLightID )
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 150


	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetPaycheckRecordset"
	cmdTemp.CommandType = adCmdStoredProc

	cmdTemp.Parameters.Refresh
	cmdTemp.Parameters("@PaycheckID") = intHighLightID
	cmdTemp.Parameters("@CustomerID") = 0


	rsPulldown.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing

	if rsPulldown.EOF then
		Set rsPulldown = Nothing
		Exit Sub
	end if

	Set ID = rsPulldown("ID")
	Set Total = rsPulldown("Total")
	Set Description = rsPulldown("Description")

	%><select name="PaycheckID" size="1"><%

	do until rsPulldown.EOF
		'Highlight the current category
		if intHighLightID = ID then
			Response.Write "<option value = '" & ID & "' SELECTED>ID #" & ID & " - " & FormatCurrency(Total) & " - " & Description & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & ID & "'>ID #" & ID & " - " & FormatCurrency(Total) & " - " & Description & "</option>" & vbCrlf
		end if

		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

End Sub
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->