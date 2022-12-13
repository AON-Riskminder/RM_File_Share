<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="file_download.aspx.vb" Inherits="RM_File_Share.file_download" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
			<table cellspacing="0" cellpadding="0" style="width: 500px; border: 1px solid #5a5a5a; text-align: left; z-index: -2; background-color: #fbfafa;">
				<tr style="background-image: url('images/menubg2.png'); background-repeat: repeat-x;">
					<td style=" border-bottom: 1px solid #5a5a5a;">
						<table cellspacing="0" cellpadding="0">
							<tr>
								<td rowspan="2" style="padding-left: 8px;">
									<asp:ImageButton runat="server" ID="ibtnDesktop" Visible="true" ImageUrl="images/riskminder.png" ToolTip="Riskminder - Optica A/S" Enabled="false" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td style="text-align: center; font-size: 14px;">
						<asp:Label ID="lblText" runat="server" Text=""></asp:Label>
					</td>
				</tr>
			</table>
		</div>
    </form>
</body>
</html>
