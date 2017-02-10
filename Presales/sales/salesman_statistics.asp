<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=salesman_statistics.asp")
%>

<p align="center"><span class=Heading>View Your Sales Statistics</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<%
'-----------------------Begin Code----------------------------
intSalesmanID = Session("EmployeeID")


Query = "SELECT * FROM Customers WHERE SalesmanID = " & intSalesmanID
Set rsMember = Server.CreateObject("ADODB.Recordset")
rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
if rsMember.EOF then
	intNumberCusts = 0
else
	intNumberCusts = rsMember.RecordCount
end if
%>
<p>You currently have <%=intNumberCusts%> customers.</p>


<%
rsMember.Close
Set rsMember = Nothing
%>


<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
