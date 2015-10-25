<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="Excel导入导出.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
   <table class="Text" cellSpacing="1" cellPadding="0" width="100%" bgColor="#1d82d0" border="0">
  <tr bgColor="#ffffff">
    <td vAlign="top">
      <table class="Text" cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tr>
          <td width="15">&nbsp;</td>
          <td vAlign="top" width="100%">
            <table class="Text" cellSpacing="1" cellPadding="0" width="100%" border="0">
              <tr height="30">
                <td style="WIDTH: 120px" width="120"><FONT face="宋体">请选择要导入的文件</FONT></td>
                <td style="WIDTH: 350px" align="left" width="350"><INPUT id="FileExcel" style="WIDTH: 300px" type="file" size="42" name="FilePhoto" runat="server"><FONT color="red"></FONT></td>
                <td class="hint"><FONT face="宋体"><asp:button id="BtnImport" Text="导 入" CssClass="button" Runat="server"></asp:button>&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="导出" />
&nbsp;&nbsp; </FONT></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <asp:label id="LblMessage" runat="server" Font-Bold="True" ForeColor="Red"></asp:label>
    </td>
  </tr>
</table>
    </form>
</body>
</html>
