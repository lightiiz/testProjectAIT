<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm2.aspx.cs" Inherits="testproject.WebForm2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h3>Export Gridview</h3>
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false" CellPadding="19">
                <Columns>
                    <asp:BoundField DataField="หมายเลขวงจร" HeaderText="หมายเลขวงจร" />
                    <asp:BoundField DataField="หน่วยงานผู้ใช้" HeaderText="หน่วยงานผู้ใช้"  />
                    <asp:BoundField DataField="SLA" HeaderText="SLA"  />
                    <asp:BoundField DataField="วันที่แจ้งเหตุ" HeaderText="วันที่แจ้งเหตุ"  />
                    <asp:BoundField DataField="เวลาแจ้งเหตุ" HeaderText="เวลาแจ้งเหตุ"  />
                    <asp:BoundField DataField="ประเภทเหตุขัดข้อง" HeaderText="ประเภทเหตุขัดข้อง" />
                    <asp:BoundField DataField="สาเหตุ" HeaderText="สาเหตุ" />
                    <asp:BoundField DataField="การแก้ไข" HeaderText="การแก้ไข"/>
                    <asp:BoundField DataField="วันที่แก้ไข" HeaderText="วันที่แก้ไข"  />
                    <asp:BoundField DataField="เวลาแก้ไข" HeaderText="เวลาแก้ไข"  />
                    <asp:BoundField DataField="ระยะเวลาการแก้ไข" HeaderText="ระยะเวลาการแก้ไข"/>
                    <asp:BoundField DataField="OS Ticket Number" HeaderText="OS Ticket Number"/>
                    <asp:BoundField DataField="หมายเหตุ" HeaderText="หมายเหตุ"  />
                    <asp:BoundField DataField="Root Cause" HeaderText="Root Cause"  />
                    <asp:BoundField DataField="Hardware" HeaderText="Hardware"  />
                    <asp:BoundField DataField="ปัญหาที่เกิด" HeaderText="ปัญหาที่เกิด" />
                    <asp:BoundField DataField="Breach/Meet" HeaderText="Breach/Meet"  />
                    <asp:BoundField DataField="ข้อยกเว้น" HeaderText="ข้อยกเว้น"/>
                    <asp:BoundField DataField="วิเคราะห์ Customer" HeaderText="วิเคราะห์ Customer"  />
                </Columns>


            </asp:GridView>
        </div>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Button" />
    </form>
</body>
</html>
