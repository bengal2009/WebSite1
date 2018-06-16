<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="StyleSheet1.css" rel="stylesheet" />
    <style type="text/css">
        #Submit1 {
            height: 19px;
            width: 84px;
        }
        .auto-style1 {
            width: 166px;
        }
    </style>
</head>
<body>
    <form id="form1" action="result.aspx" runat="server">
    <div class="testbox">
    11111111111
        <ul class="scroll-list">
            <li class="scroll-item">11111111</li>
             <li class="scroll-item">
                 <a href="Default.aspx">
                     <span class="category">11111111</span>
                 </a>
             </li>
        </ul>
    </div>
         <div class="testbox">

             <asp:Calendar ID="Calendar1" runat="server" BackColor="#FFFFCC" BorderColor="#FFCC66" BorderWidth="1px" DayNameFormat="Shortest" Font-Names="Verdana" Font-Size="8pt" ForeColor="#663399" Height="200px" ShowGridLines="True" Width="220px">
                 <DayHeaderStyle BackColor="#FFCC66" Font-Bold="True" Height="1px" />
                 <NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC" />
                 <OtherMonthDayStyle ForeColor="#CC9966" />
                 <SelectedDayStyle BackColor="#CCCCFF" Font-Bold="True" />
                 <SelectorStyle BackColor="#FFCC66" />
                 <TitleStyle BackColor="#990000" Font-Bold="True" Font-Size="9pt" ForeColor="#FFFFCC" />
                 <TodayDayStyle BackColor="#FFCC66" ForeColor="White" />
             </asp:Calendar>

             </div>
        <br />
        <div class="testbox2">
            <table><tr><td>Name:</td><td class="auto-style1">
                <input id="name1" type="text" /></td></tr>
                <tr><td>password:</td><td class="auto-style1">
                <input id="PWD" type="password" /></td></tr>
                <tr><td>&nbsp;</td><td align="center" class="auto-style1">
                    <input id="Submit1" align="middle" lang="zh-tw" size="10px" type="submit" value="送出" /></td></tr>

            </table>
        </div>
    </form>
</body>
</html>
