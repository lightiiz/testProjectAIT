<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="teststring.aspx.cs" Inherits="testproject.teststring" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
<style type="text/css">
        .auto-style1 {
            width: 46%;
            height: 124px;
        }
        .auto-style2 {
            width: 368px;
        }
        .auto-style3 {
            width: 368px;
            height: 56px;
        }
        .auto-style4 {
            height: 56px;
            width: 299px;
        }
        .auto-style5 {
            width: 299px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Import Data<br />
            <table class="auto-style1">
                <tr>
                    <td class="auto-style3">Upload Excel File</td>
                    <td class="auto-style4">
                        <asp:FileUpload ID="FileUpload1" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="auto-style2">
                        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
                    </td>
                    <td class="auto-style5">
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Upload and Save to SQL Server" />
                    </td>
                </tr>
            </table>


        </div>
    </form>
</body>
</html>