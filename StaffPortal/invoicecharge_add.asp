<!-- #include file="dsn.asp" -->



<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add Invoice Charge</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("InvoiceID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the InvoiceID."))
intInvoiceID = CInt(Request("InvoiceID"))

strSubmit = Request("Submit")

if strSubmit = "Add" then
	Set cmdReviews = Server.CreateObject("ADODB.Command")
	With cmdReviews
		.ActiveConnection = Connect
		.CommandText = "AddCustomerInvoiceCharge"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@InvoiceID") = intInvoiceID
		.Parameters("@Total") = cDbl(Request("Total"))
		if Request("Hours") <> "" then .Parameters("@Hours") = cDbl(Request("Hours"))
		.Parameters("@Description") = Format(Request("Description"))
		.Parameters("@CustomerNote") = Format(Request("CustomerNote"))
		.Parameters("@StaffNote") = Format(Request("StaffNote"))

		'Custom time included
		DateStarted = AssembleDate("DateStarted")

		DateEnded = AssembleDate("DateEnded")

		.Parameters("@DateStarted") = DateStarted
		.Parameters("@DateEnded") = DateEnded

		.Execute , , adExecuteNoRecords
	End With

	Set cmdReviews = Nothing

	currTemp = UpdateInvoicePrice( intInvoiceID )
%>
<p>The invoice charge has been added.  
<a href="invoicecharge_add.asp?InvoiceID=<%=intInvoiceID%>">Add another charge to this invoice.</a><br>
<a href="invoices_modify.asp?Submit=Edit&ID=<%=intInvoiceID%>">View the invoice.</a><br>
<a href="invoices_modify.asp">Browse the list of invoices.</a>
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
			PutDate(document.MyForm, 'DateStarted' );

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

			//1800 secs every half hour.  get the whole number of half hours (round up)
			var unroundedHalfHour = (Secs/1800);			
			var roundedHalfHour = Math.floor(unroundedHalfHour);
			//if there is a part of a half hour left, increment the number
			if ( unroundedHalfHour > roundedHalfHour )
				roundedHalfHour ++;

			varHours = (roundedHalfHour/2);

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

	//-->
	</SCRIPT>
	* indicates required information<br>
	<form method="post" action="invoicecharge_add.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="InvoiceID" value="<%=intInvoiceID%>">

	<input type="button" value="Calculate Cost" onClick="CalculateCost()">&nbsp;&nbsp;&nbsp;
	<input type="button" value="Start" onClick="StartWork()">&nbsp;&nbsp;&nbsp;
	<input type="button" value="Stop" onClick="StopWork()">

	<% PrintTableHeader 0 %>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">* Charge Total</td>
			<td class="<% PrintTDMain %>"> 
				<input type="text" name="Total" value="$" size="5">&nbsp;&nbsp;&nbsp;@ <input type="text" name="Rate" value="$" size="5"> per hour.
			</td>
   		</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Work Started</td>
      	<td class="<% PrintTDMain %>"><% DatePulldown "DateStarted", "", 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateStarted')">

		</td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Work Ended</td>
      	<td class="<% PrintTDMain %>"><% DatePulldown "DateEnded", "", 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateEnded')">


		</td>
    </tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Hours - 30 min intervals</td>
			<td class="<% PrintTDMain %>"> 
				<input type="text" name="Hours" value="" size="5">
			</td>
   		</tr>


		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Description</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Note to Customer</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="CustomerNote" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Note to Staff Only</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="StaffNote" cols="55" rows="2" wrap="PHYSICAL"></textarea>
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