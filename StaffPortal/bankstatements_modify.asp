<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Bank Statements</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankstatements_modify.asp&ID=" & Request("ID") & "&Submit=" & Request("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

strSubmit = Request("Submit")

if strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, FileName FROM BankStatements WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strPath = GetPath ("posts")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
		if FileSystem.FileExists( strPath & rsUpdate("FileName") ) then FileSystem.DeleteFile( strPath & rsUpdate("FileName") )
	Set FileSystem = Nothing

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The statement has been deleted. &nbsp;<a href="bankstatements_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM BankStatements WHERE ID = " & intID
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	rsAccount.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsAccount.EOF then
		set rsAccount = Nothing
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

	<form enctype="multipart/form-data" method="post" action="bankstatements_modify_process.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=intID%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bank Account
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintAccountsPullDown rsAccount("AccountID"), "AccountID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Started
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateStarted", rsAccount("DateStarted"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Ended
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateEnded", rsAccount("DateEnded"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Starting Balance
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="StartingBalance" value="<%=FormatCurrency(rsAccount("StartingBalance"))%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Ending Balance
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="EndingBalance" value="<%=FormatCurrency(rsAccount("EndingBalance"))%>">
			</td>
		</tr>
<%
	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	blFile = False

	if rsAccount("FileName") <> "" then
		blFile = FileSystem.FileExists( GetPath("posts") & rsAccount("FileName") )
	end if

	if blFile then 
%>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">
			Should you keep the current file? 	<% PrintRadio 1, "UseFile" %><br>
			If you want to replace the current file with a new one, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>
<%
	else
%>	
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If it is stored on a file, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>

<%
	end if


	Set FileSystem = Nothing
%>
		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note</td>
     		<td class="<% PrintTDMain %>"> 
    			<textarea name="Note" cols="55" rows="4" wrap="PHYSICAL"><%=FormatEdit(rsAccount("Note"))%></textarea>
    		</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>
	</table>
	</form>



<%
'-----------------------Begin Code----------------------------
	rsAccount.Close
	Set rsAccount = Nothing

else
	Query = "SELECT * FROM BankStatements ORDER BY DateEnded DESC, ID DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You have to have statements before you can modify them."))
	end if

	Set ID = rsPage("ID")
	Set DateStarted = rsPage("DateStarted")
	Set DateEnded = rsPage("DateEnded")
	Set Note = rsPage("Note")
	Set AccountID = rsPage("AccountID")
	Set StartingBalance = rsPage("StartingBalance")
	Set EndingBalance = rsPage("EndingBalance")
	Set FileName = rsPage("FileName")


	strPath = GetPath ("posts")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
'-----------------------End Code----------------------------
%>
	<p><a href="bankstatements_add.asp">Add A Statement</a></p>
	<form METHOD="POST" ACTION="bankstatements_modify.asp">
<%
'-----------------------Begin Code----------------------------
	PrintPagesHeader
'-----------------------End Code----------------------------
%>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">Account</td>
		<td class="TDHeader">Statement Period</td>
		<td class="TDHeader">Balances</td>
		<td class="TDHeader">Difference</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for p = 1 to rsPage.PageSize
		if not rsPage.EOF then
			blExists = FileSystem.FileExists( strPath & rsPage("FileName") )
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="bankstatements_modify.asp">
		<input type="hidden" name="ID" value="<%=ID%>">
			<tr>
				<td class="<% PrintTDMain %>">
<%
				if blExists then
					%><a href="posts/<%=FileName%>">View</a><%
				else
					Response.Write "&nbsp;"
				end if
%>

				</td>
				<td class="<% PrintTDMain %>"><%=GetAccountName( AccountID )%></td>
				<td class="<% PrintTDMain %>"><%=FormatDateTime(DateStarted, 2)%> - <%=FormatDateTime(DateEnded, 2)%></td>
				<td class="<% PrintTDMain %>"><%=FormatCurrency(StartingBalance)%> - <%=FormatCurrency(EndingBalance)%></td>
				<td class="<% PrintTDMain %>"><%=FormatCurrency(cDbl(EndingBalance) - cDbl(StartingBalance))%></td>
				<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this account, there is no way to get it back.  Are you sure?', 'bankstatements_modify.asp?Submit=Delete&ID=<%=ID%>')">				
				</td>
			</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
		rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
	rsPage.Close
	set rsPage = Nothing

	Set FileSystem = Nothing

end if

'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->