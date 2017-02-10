<%
Sub PrintBullet
	%><img src="../images/bullet.gif" border="0"><%
End Sub

Sub PrintArrow
	%><img src="../images/arrow.gif" border="0"><%
End Sub


Sub PrintSymb( strType, strImage )
	strType = lcase(strType)
	strImage = lcase(strImage)
	if strType = "home" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "member" or strType = "members" then
		strType = "members"
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "basicpost" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "browse" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>e.gif" border="0"></a><%
		end if
	elseif strType = "create" then
		strType = "create"
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "confirmation" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "delete" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "edit" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "list" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "login" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "popup" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "property" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "search" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "viewing" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	elseif strType = "warning" then
		if strImage = "" then
			%><a href="javascript:openWindow('screenshots/<%=strType%>.gif');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		elseif strImage = "none" then
			%><img src="../images/symbol_<%=strType%>.gif" border="0"><%
		else
			%><a href="javascript:openWindow('screenshots/<%=strImage%>');"><img src="../images/symbol_<%=strType%>.gif" border="0"></a><%
		end if
	else
		%><span class="SubHeading"><%=strType%></span><%
	end if
End Sub
%>