<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Statement</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankstatements_add.asp")
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

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

			form.StartingBalance.value = ConvertDollar(form.StartingBalance.value);
			form.EndingBalance.value = ConvertDollar(form.EndingBalance.value);

			if (form.StartingBalance.value == "" )
				strError += "          You forgot the starting balance. \n";
			if (form.EndingBalance.value == "" )
				strError += "          You forgot the ending balance. \n";

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

	<form enctype="multipart/form-data" method="post" action="bankstatements_add_process.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bank Account
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintAccountsPullDown 0, "AccountID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Started
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateStarted", Date, 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Ended
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateEnded", Date, 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Starting Balance
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="StartingBalance" value="$">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Ending Balance
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="EndingBalance" value="$">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If it is stored on a file, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>

		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note</td>
     		<td class="<% PrintTDMain %>"> 
    			<textarea name="Note" cols="55" rows="4" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Add">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------


'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->