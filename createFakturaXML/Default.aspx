<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="createFakturaXML._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="height: 239px">
        <asp:FileUpload ID="FileUpload1" runat="server" Width="802px"/>        
        <asp:Button ID="Button3" runat="server" style="margin-left: 631px" Text="Сформировать фаил" Width="173px" OnClick="Button3_Click"/>
        <br />
        <asp:CheckBox ID="isTlt" runat="server" Text="Тольятти" /><asp:CheckBox ID="isProseka" runat="server" Text="7 просека" /><asp:CheckBox ID="isRaduga" runat="server" Text="Радужный элит" />
        <br />
        <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
        <asp:Button ID="Button4" runat="server" OnClick="Button4_Click" Text="Button" />
    </div>
    </form>
</body>
</html>
