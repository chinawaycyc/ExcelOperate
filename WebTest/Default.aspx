<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="btnExport" runat="server" Text="导出" OnClick="btnExport_Click" />
        <asp:FileUpload ID="fileUpload" runat="server" />
        <asp:Button ID="btnImport" runat="server" Text="导入" OnClick="btnImport_Click" />
        <asp:GridView ID="gvPS" runat="server" AutoGenerateColumns="False">
            <Columns>
                <asp:BoundField DataField="序号" HeaderText="序号" />
                <asp:BoundField DataField="姓名" HeaderText="姓名" />
                <asp:BoundField DataField="性别" HeaderText="性别" />
                <asp:BoundField DataField="身份证" HeaderText="身份证" />
                <asp:BoundField DataField="随机唯一标识码" HeaderText="随机唯一标识码" />
            </Columns>
        </asp:GridView>
    </div>
    </form>
</body>
</html>
