<%
'-----------------------Begin Code----------------------------
Table = Request("Table")

intCategoryID = Request("ID")
if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)

if not ValidCategory(intCategoryID, Table) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid category."))

Response.Write GetCatHeiarchy( intCategoryID, "store.asp", Table, StoreTitle )

GetCategoryInfo intCategoryID, strName, strBody

'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading><%=strName%></span><br>
<a href="javascript:history.go(-1)">Back</a><br>
<a href="store_cart_view.asp?CartID=<%=GetCartID()%>">My Cart</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="store_checkout.asp?CartID=<%=GetCartID()%>">Checkout</a>
</span></p>
<%
Response.Write strBody
%>
