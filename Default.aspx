<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Access Token:
            <asp:TextBox Text="" ID="tbAccessToken" runat="server" Width ="200px"/>
            <asp:Button Text="Connect" ID="btnOAuthConnect" runat="server" OnClick="btnOAuthConnect_Click"/>
        </div>
        
        <fieldset style="width: 800px">
         <div>
            Search products by:
            <asp:DropDownList ID="productSearchDD" runat="server">
                <asp:ListItem Value="Country" Selected="True">
                    Country
                </asp:ListItem>
                <asp:ListItem Value="Category" Selected="False">
                    Category
                </asp:ListItem>
            </asp:DropDownList>
             <asp:TextBox Text="" ID="tbSearch" runat="server" Width="200px"/>
             <asp:Button Text="Search" ID="btnProductSearch" runat="server" OnClick="btnProductSearch_Click" />
         </div>
         <div>
             <asp:Button Text="View Average order price per country" ID="btnAvgPrice" runat="server" OnClick="btnAvgPrice_Click"/>
         </div>
        </fieldset>
        <asp:GridView ID="gvSheets" runat="server" AutoGenerateColumns="True" AllowSorting="true" OnSorting="gvSheets_Sorting"/>
   </form>
</body>
</html>
