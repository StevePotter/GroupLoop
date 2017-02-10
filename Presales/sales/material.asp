<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=material.asp")
%>

<p align="center"><span class=Heading>Material to Distribute</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<p align="center" class=SubHeading>
Our Flier
</p>
<p>
Print out this flier to give to each church you visit.<br>
<b>BE SURE TO WRITE DOWN YOUR SALESMAN NUMBER (<%=Session("EmployeeID")%>) ON EACH PRINTOUT!</b>  <br>
This is <i>extremely</i> important.  Make it clear that if they sign up, they need to enter your salesman number.  This way we know you are the one 
to give the money to.
</p>

<p>
<a href="flier2.doc">Download the Flier for Any Organization</a> (Microsoft Word)<br>
<a href="flier2.pdf">Download the Flier for Any Organization</a> (Adobe Acrobat)<br><br>

<a href="flier.doc">Download the Flier for Churches</a> (Microsoft Word)<br>
<a href="flier.pdf">Download the Flier for Churches</a> (Adobe Acrobat)

</p>

<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
