<script language="VB" runat="server">
	public mySection As Section


	mySection.AddItem( tblForm, lblSubmitPage )


	sub Submit( obj as object, e as eventargs )
		pnlForm.Visible = false
		pnlSubmitted.Visible = true
		hlAgain.NavigateURL = mySection.Links("Add") & "?ID=" & mySection.ID
		mySection.Add( tblForm )


'		Response.Write("-" & mlCtrl.Text & "-")
	end sub
</script>


		<asp:Panel id="pnlForm" runat="server">
			<form runat="server" OnSubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">

				<script language="JavaScript">
				<!--
					function submit_page(form) {
						//Error message variable
						var strError = "";
						<asp:Literal id="lblSubmitPage" runat=server />

						if (strError == "") {
							return true;
						}
						else{
							strError = "Sorry, but you must go back and fix the following errors: \n" + strError;
							alert (strError);
							return false;
						}   
					}

				//-->
				</script>


			<asp:Table id="tblForm" runat="server" >

				<asp:TableRow runat="server">
					<asp:TableCell runat="server" ColumnSpan="2" cssClass="TDMain1" HorizontalAlign="center">
						<asp:Button runat="server" Text="Add" OnClick="Submit" />
					</asp:TableCell>
				</asp:TableRow>
			</asp:Table>
			</form>
		</asp:Panel>

		<asp:Panel id="pnlSubmitted" runat="server" visible="false">
			The <%=mySection.Display("NounSingular")%> has been added.  <asp:HyperLink id="hlAgain" runat="server">Add another.</asp:HyperLink>
		</asp:Panel>