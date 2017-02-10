<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
strSource = Request("Source")
if strSource = "" then strSource = "index.asp"

Response.Cookies("NoWYSIWYGEdit").expires = #10/10/2020#
Response.Cookies("NoWYSIWYGEdit").path = "/"

'they indicated to not use activedit
if Request("NoUseWYSIWYGEdit") = "1" then
	Response.Cookies("NoWYSIWYGEdit") = "1"

	Response.write "You are now not using the advanced editor.  Please click Back on your browser, and if it still appears, click Refresh."

'they indicated to use it
elseif Request("UseWYSIWYGEdit") = "1" then
	Response.Cookies("NoWYSIWYGEdit") = ""
	Response.write "You are now using the advanced editor.  Please click Back on your browser, and if it does not appear, click Refresh."
end if

	'if InStr( strSource, "?") then
	'	Response.Redirect(strSource & "&ID=" & Request("ID") )
'	elseif Request("ID") <> "" then
	'	Response.Redirect(strSource & "?ID=" & Request("ID") )
'	else
	'	Response.Redirect(strSource)
'	end if

'end if



'------------------------End Code-----------------------------
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->