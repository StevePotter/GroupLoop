<html>
<head>
<title>GroupLoop.com Sales Program</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style type="text/css">
BODY{font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; color: #000000}
.BodyText{font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; color: #000000}
.LinkText{font-family:Verdana, Arial, Helvetica, sans-serif; font-size:14px; color: #FFFF66}
.Heading{font-family:Verdana, Arial, Helvetica, sans-serif; font-size:20px; color: #49838D; font-weight: bold ; font-style: italic}
.SubHeading{font-family:"Times New Roman", Times, serif; font-size:18px; color: #49838D; font-weight: bold; font-style: italic}
.Title{font-family:"Times New Roman", Times, serif; font-size:30px; color: #000000; font-weight: bold; font-style: normal}
.TDMain1{font-family:Arial, Helvetica, sans-serif; font-size:10px; color: #000000; background-color: #99FFFF }
.TDMain2{font-family:Arial, Helvetica, sans-serif; font-size:10px; color: #000000; background-color: #66FFFF }
.TDHeader{font-family:Arial, Helvetica, sans-serif; font-size:14px; color: #000000; background-color: #66CCFF ; font-weight: bold ; font-style: italic}
</style>

<script language="JavaScript1.2" src="..\scripts.js" type="text/javascript"></script>

</head>
<%
CellSpacing = 1
CellPadding = 2
Border = 0

Function GetReferralLink()
	if Session("ReferralID") <> "" then
		GetReferralLink = "ReferralID=" & Session("ReferralID")
	elseif Request("ReferralID") <> "" then
		Session("ReferralID") = CInt(Request("ReferralID"))
		GetReferralLink = "ReferralID=" & Session("ReferralID")
	else
		GetReferralLink = ""
	end if

End Function
%>
<body bgcolor="#FFFFFF" link="#2BACAC" vlink="1E7575" alink="#2BACAC" onLoad="<%=strOnLoad%>">

<table width="100%" border="0" cellspacing="0" cellpadding="0" border="0">
	<tr>
	<td width="100%" valign="top" align="center">
		<a href="http://www.grouploop.com/sales/index.asp?<%=GetReferralLink()%>"><img src="title.gif" border=0></a><br><br>
	</td>
	</tr>
	<tr>
	<td width="100%" valign="top" align="left" class="BodyText">
