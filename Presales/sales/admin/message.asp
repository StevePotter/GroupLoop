<!-- #include file="header.asp" -->

<br>
<%
'
'-----------------------Begin Code----------------------------
'This just displays a message that is passed.

strTitle = Request("Title")
strSource = Request("Source")
strMessage = Request("Message")

%>
<p align=center class=Heading><%=strTitle%></p>
<%

'Display the message if there is one
if strMessage <> "" then
%>	<b><%=strMessage%></b><br>	<%
end if

'Don't give the back option
if LCase(strSource) = "noback" then
	Response.Write ""
'Give the given link
elseif strSource <> "" then
	%>	<a href="<%=strSource%>">Click here</a> to go back.<%
else
	%>	<a href="javascript:history.back(1)">Click here</a> to go back.<%
end if
'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->
