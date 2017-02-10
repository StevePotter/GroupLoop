<!-- #include file="admin_functions.asp" -->
<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin then	Redirect("members.asp?Source=admin_look_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Colors and Fonts</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Update" then
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic



	rsLook("BackgroundColor") = GetColor("BackgroundColor")

	rsLook("BodyTextSize") = Request("BodyTextSize")
	rsLook("BodyTextColor") = GetColor("BodyTextColor")
	rsLook("BodyTextBold") = Request("BodyTextBold")
	rsLook("BodyTextFont") = Request("BodyTextFont")
	rsLook("BodyTextItalic") = Request("BodyTextItalic")


	rsLook("TableMainBackground1") = GetColor("TableMainBackground1")
	rsLook("TableMainBackground2") = GetColor("TableMainBackground2")
	rsLook("TableMainTextColor") = GetColor("TableMainTextColor")
	rsLook("TableMainTextSize") = Request("TableMainTextSize")
	rsLook("TableMainTextBold") = Request("TableMainTextBold")
	rsLook("TableMainTextFont") = Request("TableMainTextFont")
	rsLook("TableMainTextItalic") = Request("TableMainTextItalic")

	rsLook("TableHeaderBackground") = GetColor("TableHeaderBackground")
	rsLook("TableHeaderTextColor") = GetColor("TableHeaderTextColor")
	rsLook("TableHeaderTextSize") = Request("TableHeaderTextSize")
	rsLook("TableHeaderTextBold") = Request("TableHeaderTextBold")
	rsLook("TableHeaderTextItalic") = Request("TableHeaderTextItalic")
	rsLook("TableHeaderTextFont") = Request("TableHeaderTextFont")

	rsLook("HeadingColor") = GetColor("HeadingColor")
	rsLook("HeadingSize") = Request("HeadingSize")
	rsLook("HeadingBold") = Request("HeadingBold")
	rsLook("HeadingItalic") = Request("HeadingItalic")
	rsLook("HeadingFont") = Request("HeadingFont")

	rsLook("TitleColor") = GetColor("TitleColor")
	rsLook("TitleSize") = Request("TitleSize")

	rsLook("TitleBold") = Request("TitleBold")
	rsLook("TitleItalic") = Request("TitleItalic")
	rsLook("TitleFont") = Request("TitleFont")


	rsLook("LinkColor") = GetColor("LinkColor")
	rsLook("VisitedLinkColor") = GetColor("VisitedLinkColor")
	rsLook("LinkSize") = Request("LinkSize")

	if Request("LeftMenuSize") <> "" then
		rsLook("LeftMenuColor") = GetColor("LeftMenuColor")
		rsLook("LeftMenuSize") = Request("LeftMenuSize")
		rsLook("LeftMenuBold") = Request("LeftMenuBold")
		rsLook("LeftMenuItalic") = Request("LeftMenuItalic")
		rsLook("LeftMenuUnderline") = Request("LeftMenuUnderline")
		rsLook("LeftMenuFont") = Request("LeftMenuFont")
	end if

	if Request("RightMenuSize") <> "" then
		rsLook("RightMenuColor") = GetColor("RightMenuColor")
		rsLook("RightMenuSize") = Request("RightMenuSize")
		rsLook("RightMenuBold") = Request("RightMenuBold")
		rsLook("RightMenuUnderline") = Request("RightMenuUnderline")
		rsLook("RightMenuItalic") = Request("RightMenuItalic")
		rsLook("RightMenuFont") = Request("RightMenuFont")
	end if

	if Request("TopMenuSize") <> "" then
		rsLook("TopMenuColor") = GetColor("TopMenuColor")
		rsLook("TopMenuSize") = Request("TopMenuSize")
		rsLook("TopMenuBold") = Request("TopMenuBold")
		rsLook("TopMenuUnderline") = Request("TopMenuUnderline")
		rsLook("TopMenuItalic") = Request("TopMenuItalic")
		rsLook("TopMenuFont") = Request("TopMenuFont")
	end if


	rsLook.Update
	set rsLook = Nothing
%>
	<!-- #include file="write_constants.asp" -->
<%
	Redirect("write_header_footer.asp?Source=admin_look_edit.asp?Submit=Changed")

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The changes have been made.  You can <a href="admin_look_edit.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>
<%
'-----------------------Begin Code---------------------b-------
else
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly

	Set rsItems = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT Name FROM Fonts ORDER BY ID"
	rsItems.CacheSize = 5
	rsItems.Open Query, Connect, adOpenStatic, adLockReadOnly
	Set Name = rsItems("Name")
	Dim strFonts(10, 2)
	intMaxFonts = rsItems.RecordCount
	for i = 1 to intMaxFonts
		strFonts(i, 0) = Name
		MyArray = Split(Name, ",", -1, 1)
		strFonts(i, 1) = MyArray(0)
		rsItems.MoveNext
	next
	rsItems.Close
	set rsItems = Nothing


	Sub PrintIcon( strFieldName, strType, intFieldValue )
		intFieldValue = cInt(intFieldValue)

		strOnOff = "off"
		if intFieldValue = 1 then strOnOff = "on"

		'Hidden field needed for the thing
		'and also the link to the function
%>
		<input type="hidden" name="<%=strFieldName%>" value="<%=intFieldValue%>">
		<a href="javascript:ChangeHidden('<%=strFieldName%>', '<%=strType%>');"><img src="http://www.GroupLoop.com/homegroup/images/<%=strType%>_button_<%=strOnOff%>.gif" border="0" name="<%=strFieldName%>Img"></a>
<%
	End Sub

%>
	<a href="javascript:(alert('The reason you can\'t see the colors in the menus is because your Internet browser is not current.  Updates are free, and you can visit www.Netscape.com or www.Microsoft.com to update.'))">Can't see colors in menus?<br>

	<script language="JavaScript">
	<!--

	//This changes the icons for the bold and italic
	function ChangeHidden( Name, Type )
	{	
		var Field, Img;

		Field = document.myForm.elements[Name];
		Img = document.images[Name+'Img'];

		//it is not bold right now, so make it bold
		if (Field.value == '0'){
			Field.value = '1';
			Img.src='http://www.GroupLoop.com/homegroup/images/' + Type + '_button_on.gif'
		}else{
			//Make it unbold
			Field.value = '0';
			Img.src='http://www.GroupLoop.com/homegroup/images/' + Type + '_button_off.gif'

		}
	}

	//This will allow people to enter a custom color
	function ChangeHidden( Name, Type )
	{
		var Field, Img;

		Field = document.myForm.elements[Name];
		Img = document.images[Name+'Img'];

		//it is not bold right now, so make it bold
		if (Field.value == '0'){
			Field.value = '1';
			Img.src='http://www.GroupLoop.com/homegroup/images/' + Type + '_button_on.gif'
		}else{
			//Make it unbold
			Field.value = '0';
			Img.src='http://www.GroupLoop.com/homegroup/images/' + Type + '_button_off.gif'

		}
	}


function Launch(Field) {
color=open("","newWindow","resizable=1,width=450,height=140,top=0,left=0");
color.document.write('<html><head><title>Enter Your Custom Color</title></head><body><p align=center>If you are an advanced user and wish to use a custom color, please enter it below.  Colors are recommended to be in the #AABBCC format.  If you do not know what you are doing, just close this box and choose a different color.</p><p align=center><form name=newForm onSubmit="window.opener.document.myForm.'+Field+'.value=document.newForm.NewColor.value; window.close(); return false;"><input type="Text" name="NewColor" size=8><input type=submit name=Submit value="Use This Color">\n</form></p></BODY></HTML>\n');
}//end colors


//-->
	</SCRIPT>



	<form METHOD="post" ACTION="admin_look_edit.asp" name="myForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">

 


<%
	PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Page Element
			</td>
		</tr>
<%
	'If they have an image covering the title, don't confuse them with this
	if not ImageExists( "TitleImage", strExt) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Site Title (at top of each page)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "TitleFont", rsLook("TitleFont")  %>
				<% PrintSizePullDown "TitleSize", rsLook("TitleSize" ) %>
				<% PrintIcon "TitleBold", "bold", rsLook("TitleBold")  %>
				<% PrintIcon "TitleItalic", "italic", rsLook("TitleItalic")  %>
				<% PrintArrayColors strColors, intMaxColors,  "TitleColor", rsLook("TitleColor")%>
			</td>
		</tr>
<%
	end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Main Headings (such as the 'Change Colors and Fonts' heading above)
			</td>
			<td class="<% PrintTDMain %>" align="left">
<%'				<input type="text" name="Test" value="" onchange="Tester();">%>
				<% PrintFontsPullDown "HeadingFont", rsLook("HeadingFont")  %>
				<% PrintSizePullDown "HeadingSize", rsLook("HeadingSize" ) %>
				<% PrintIcon "HeadingBold", "bold", rsLook("HeadingBold")  %>
				<% PrintIcon "HeadingItalic", "italic", rsLook("HeadingItalic")  %>
				<% PrintArrayColors strColors, intMaxColors,  "HeadingColor", rsLook("HeadingColor")%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Regular Body Text (such as the body of a story, announcement, etc)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "BodyTextFont", rsLook("BodyTextFont")  %>
				<% PrintSizePullDown "BodyTextSize", rsLook("BodyTextSize" ) %>
				<% PrintIcon "BodyTextBold", "bold", rsLook("BodyTextBold")  %>
				<% PrintIcon "BodyTextItalic", "italic", rsLook("BodyTextItalic")  %>
				<% PrintArrayColors strColors, intMaxColors,  "BodyTextColor", rsLook("BodyTextColor")%>
			</td>
		</tr>
<%
		if SectorHasButtons("Top") then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				The buttons (links) in the top menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "TopMenuFont", rsLook("TopMenuFont")  %>
				<% PrintSizePullDown "TopMenuSize", rsLook("TopMenuSize" ) %>
				<% PrintIcon "TopMenuBold", "bold", rsLook("TopMenuBold")  %>
				<% PrintIcon "TopMenuItalic", "italic", rsLook("TopMenuItalic")  %>
				<% PrintIcon "TopMenuUnderline", "Underline", rsLook("TopMenuUnderline")  %>
				<% PrintArrayColors strColors, intMaxColors,  "TopMenuColor", rsLook("TopMenuColor")%>
			</td>
		</tr>
<%
		end if
		if SectorHasButtons("Left") then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				The buttons (links) in the left menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "LeftMenuFont", rsLook("LeftMenuFont")  %>
				<% PrintSizePullDown "LeftMenuSize", rsLook("LeftMenuSize" ) %>
				<% PrintIcon "LeftMenuBold", "bold", rsLook("LeftMenuBold")  %>
				<% PrintIcon "LeftMenuItalic", "italic", rsLook("LeftMenuItalic")  %>
				<% PrintIcon "LeftMenuUnderline", "Underline", rsLook("LeftMenuUnderline")  %>
				<% PrintArrayColors strColors, intMaxColors,  "LeftMenuColor", rsLook("LeftMenuColor")%>
			</td>
		</tr>
<%
		end if
		if SectorHasButtons("Right") then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				The buttons (links) in the right menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "RightMenuFont", rsLook("RightMenuFont")  %>
				<% PrintSizePullDown "RightMenuSize", rsLook("RightMenuSize" ) %>
				<% PrintIcon "RightMenuBold", "bold", rsLook("RightMenuBold")  %>
				<% PrintIcon "RightMenuItalic", "italic", rsLook("RightMenuItalic")  %>
				<% PrintIcon "RightMenuUnderline", "Underline", rsLook("RightMenuUnderline")  %>
				<% PrintArrayColors strColors, intMaxColors,  "RightMenuColor", rsLook("RightMenuColor")%>
			</td>
		</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Table Header Text (such as the 'Page Element' title above)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "TableHeaderTextFont", rsLook("TableHeaderTextFont")  %>
				<% PrintSizePullDown "TableHeaderTextSize", rsLook("TableHeaderTextSize" ) %>
				<% PrintIcon "TableHeaderTextBold", "bold", rsLook("TableHeaderTextBold")  %>
				<% PrintIcon "TableHeaderTextItalic", "italic", rsLook("TableHeaderTextItalic")  %>
				<% PrintArrayColors strColors, intMaxColors,  "TableHeaderTextColor", rsLook("TableHeaderTextColor")%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Table Main Text (such as this very text)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintFontsPullDown "TableMainTextFont", rsLook("TableMainTextFont")  %>
				<% PrintSizePullDown "TableMainTextSize", rsLook("TableMainTextSize" ) %>
				<% PrintIcon "TableMainTextBold", "bold", rsLook("TableMainTextBold")  %>
				<% PrintIcon "TableMainTextItalic", "italic", rsLook("TableMainTextItalic")  %>
				<% PrintArrayColors strColors, intMaxColors,  "TableMainTextColor", rsLook("TableMainTextColor")%>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Hyperlink Size and Color (such as the 'Can't see colors in menus?' link above)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintSizePullDown "LinkSize", rsLook("LinkSize" ) %>

				<% PrintArrayColors strColors, intMaxColors,  "LinkColor", rsLook("LinkColor")  %>			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Visited Hyperlink Color
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintArrayColors strColors, intMaxColors,  "VisitedLinkColor", rsLook("VisitedLinkColor")  %>
			</td>
		</tr>		

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Background Colors
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Main Background Color
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintArrayColors strColors, intMaxColors,  "BackgroundColor", rsLook("BackgroundColor")  %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Background Color For Titles In Table
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintArrayColors strColors, intMaxColors,  "TableHeaderBackground", rsLook("TableHeaderBackground")  %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				1st Background Color For Main Data In Table (1st and 2nd alternate)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintArrayColors strColors, intMaxColors,  "TableMainBackground1", rsLook("TableMainBackground1")  %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				2nd Background Color For Main Data In Table (use 1st for one color)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintArrayColors strColors, intMaxColors,  "TableMainBackground2", rsLook("TableMainBackground2")  %>
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
	set rsLook = Nothing
end if

Sub PrintArrayColors ( strColorArray(), intMax, strName, strSelectColor )

	blFound = false


		HexD = Array("00", "33", "66", "99", "CC", "FF")
%>
		<input type="hidden" name="<%=strName%>Cust" value="">
		<select name="<%=strName%>" onChange="if(this.options[this.selectedIndex].value == 'Custom') Launch('<%=strName%>Cust');">
<%
		if strSelectColor = "" then
			strSelect = " selected"
			blFound = true
		end if
%>		<option value="" <%=strSelect%>>None</option>
<%
for i= 0 to 5

temp = (i mod 6)
current = HexD(temp)


for b_i = 0 to 5

	b_temp = (b_i mod 6)
	b_current = HexD(b_temp)

	for c_i = 0 to 5

	c_temp = (c_i mod 6)
	c_current = HexD(c_temp)
	all = c_current&b_current&current
		if "#" & all = strSelectColor then
			blFound = true

%>
				<option value="#<%=all%>" style="BACKGROUND: #<%=all%>;" selected><%=all%></option>
<%
		else
%>
				<option value="#<%=all%>" style="BACKGROUND: #<%=all%>;"><%=all%></option>
<%
		end if
	next

next

next

		'They are using another color
		if blFound = false then
			strValue = strSelectColor
%>
				<option value="<%=strValue%>" style="BACKGROUND: <%=strValue%>;" selected>Custom Color: <%=strValue%></option>
				<option value="Custom">Use New Custom Color</option>
<%
		else
%>
				<option value="Custom">Use Custom</option>

<%
		end if

		Response.Write "</select>"
End Sub



'-------------------------------------------------------------
'This function writes a pulldown menu for members
'-------------------------------------------------------------
Sub PrintFontsPullDown( strName, strSelectFont )
%>
	<select name="<%=strName%>">
<%
	for f = 1 to intMaxFonts
		if strFonts(f, 0) = strSelectFont then
%>
			<option value="<%=strFonts(f, 0)%>" style="font-family:'<%=strFonts(f, 0)%>';" selected><%=strFonts(f, 1)%></option>
<%
		else
%>
			<option value="<%=strFonts(f, 0)%>" style="font-family:'<%=strFonts(f, 0)%>';"><%=strFonts(f, 1)%></option>
<%
		end if
	next
%>
	</select>
<%
End Sub


'-------------------------------------------------------------
'This function writes a pulldown menu for the different font sizes
'-------------------------------------------------------------
Sub PrintSizePullDown( strName, intSelectSize )
%>
	<select name="<%=strName%>">
<%
	if blPrintEmpty then Response.Write "<option value=''> </option>" & vbCrLf

	for i = 6 To 56 Step 2
		if i = intSelectSize then	'this is the selected one
			Response.Write "<option value='" & intSelectSize &"' selected>" & intSelectSize &"</option>" & vbCrLf
		elseif (intSelectSize = i - 1) then 'this is the selected, and print the next one cuz the answer is odd
			Response.Write "<option value='" & intSelectSize &"' selected>" & intSelectSize &"</option>" & vbCrLf
			Response.Write "<option value='" & i &"'>" & i &"</option>" & vbCrLf

		else	'Not selected
			Response.Write "<option value='" & i &"' >" & i &"</option>" & vbCrLf
		end if
	next
	%>
	</select>
<%
End Sub

'------------------------End Code-----------------------------
%>
