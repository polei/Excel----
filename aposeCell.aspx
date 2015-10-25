<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="aposeCell.aspx.cs" Inherits="Excel导入导出.aposeCell" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button ID="import" runat="server" Text="导入" OnClick="import_Click" />          
        <asp:Button ID="export" runat="server" Text="导出" OnClick="export_Click" />
        <asp:Button ID="downloadModel" runat="server" Text="导出模板" OnClick="downloadModel_Click" />
        <br />
        <asp:Label ID="info" runat="server" Text=""></asp:Label>
    </div>

        <asp:GridView ID="GridView1" runat="server"></asp:GridView>
        <asp:GridView ID="GridView2" runat="server"></asp:GridView>

    </form>
</body>
</html>
