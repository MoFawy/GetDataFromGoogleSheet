<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="GetDataFromGoogleSheet.Forms.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="Asset" DataSourceID="SqlDataSource1">
                <Columns>
                    <asp:BoundField DataField="Asset" HeaderText="Asset" ReadOnly="True" SortExpression="Asset"></asp:BoundField>
                    <asp:BoundField DataField="Asset Name" HeaderText="Asset Name" SortExpression="Asset Name"></asp:BoundField>
                    <asp:BoundField DataField="Model" HeaderText="Model" SortExpression="Model"></asp:BoundField>
                    <asp:BoundField DataField="Vendor" HeaderText="Vendor" SortExpression="Vendor"></asp:BoundField>
                    <asp:BoundField DataField="Description" HeaderText="Description" SortExpression="Description"></asp:BoundField>
                </Columns>
            </asp:GridView>
            <asp:SqlDataSource runat="server" ID="SqlDataSource1" ConnectionString='<%$ ConnectionStrings:DBspreadsheetConnectionString %>' SelectCommand="SELECT * FROM [Asset]"></asp:SqlDataSource>
        </div>
    </form>
</body>
</html>
