<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="netkastring.aspx.cs" Inherits="testproject.netkastring" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <table class="auto-style1">
                <tr>
                    <td class="auto-style3">Upload Excel File</td>
                    <td class="auto-style4">
                        <asp:FileUpload ID="FileUpload2" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="auto-style2">
                        <asp:Label ID="Label2" runat="server" Text="Label"></asp:Label>
                    </td>
                    <td class="auto-style5">
                        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="Upload and Save to SQL Server" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
