<!-- #include file="header.asp" -->

<p class="Heading" align="center">Uh-Oh</p>

<p>Sorry, but there has been a problem with what you are trying to do.</p>
<%
'
'-----------------------Begin Code----------------------------	
if Request("Message") <> "" then
%>	<p>The error reported was: <b><%=Request("Message")%></b></p>	<%
end if
if Request("Source") <> "" then
%>	<p><a href="<%=Request("Source")%>">Click here</a> to go back.</p>	<%
end if
'------------------------End Code-----------------------------
%>

<p>
If this problem keeps occuring, please email <a href="mailto:support@grouploop.com">GroupLoop.com Support</a> and 
tell it to us.  We will correct the problem right away.  Thank you.
</p>

<!-- #include file="footer.asp" -->
