<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->


<%
Query = "SELECT * FROM Look"

Set rsLook = Server.CreateObject("ADODB.Recordset")
rsLook.CacheSize = 100
rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

Query = "SELECT * FROM Schemes"
Set rsSchemes = Server.CreateObject("ADODB.Recordset")
rsSchemes.CacheSize = 100
rsSchemes.Open Query, Connect, adOpenStatic, adLockOptimistic


MaxFields = rsLook.Fields.Count
Dim FieldArray(200)

NewMax = 0
for i = 0 to MaxFields - 1
	strField = rsLook(i).Name

	blExclude = cBool( Instr(strField, "ID") or Left(strField, 8) = "InfoText" or Left(strField, 7) = "Include" or Left(strField, 8) = "ListType" or Left(strField, 7) = "Display" )

	if not blExclude then
		FieldArray(NewMax) = strField
		NewMax = NewMax + 1
	end if
next


do until rsSchemes.EOF
	rsLook.AddNew
	rsLook("SchemeID") = rsSchemes("ID")

	Response.Write rsSchemes("ID") & "<br>"

	for i = 0 to NewMax - 1
		Response.Write FieldArray(i) & "<br>"

		rsLook(FieldArray(i)) = rsSchemes(FieldArray(i))
	next

	rsLook.Update

	rsSchemes.MoveNext
loop

rsSchemes.Close
Set rsSchemes = Nothing

rsLook.Close
Set rsLook = Nothing



%>
dun

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->