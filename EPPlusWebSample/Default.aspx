<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="EPPlusWebSample._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>EPPlus Websample</title>
  <style type="text/css">
    body
    {
        BACKGROUND-COLOR:#FFFFFF;
        COLOR: #111111;
        FONT-FAMILY: Arial;
        font-size: large;
        margin-top: 1px
    }
    table
    {
        FONT-FAMILY: Arial;
        FONT-SIZE: medium
    }    
A:link {color:#003399;}
A:visited {color:#003399;}
A:hover{color:#CC6699}

   </style>
</head>
<body>
    <form id="form1" runat="server">
    <h1>EPPlus Web samples</h1>
        <h3>The web sample project shows a few different ways to send a workbook to the client. </h3>
        <table>
        <tr>
        <td>
            <asp:HyperLink ID="sample1" runat="server" NavigateUrl="~/GetSample.aspx?Sample=1">Sample 1</asp:HyperLink>
        </td>
        <td>
            This sample demonstrates how to output a spreadsheet using the SaveAs(Reponse.OutputSteam) method.
        </td>
        </tr>
        <tr>
        <td>
        <asp:HyperLink ID="sample2" runat="server" NavigateUrl="~/GetSample.aspx?Sample=2">Sample 2</asp:HyperLink>
        </td>
        <td>
            This sample demonstrates how to output a spreadsheet using the Response.BinaryWrite(pck.GetAsByteArray()).
        </td>
        </tr>
        <tr>
        <td>
        <asp:HyperLink ID="sample3" runat="server" NavigateUrl="~/GetSample.aspx?Sample=3">Sample 3</asp:HyperLink>
        </td>
        <td>
            This sample demonstrates how to use a template stored in the Application cache.
        </td>
        </tr>
        <tr>
        <td>
        <asp:HyperLink ID="sample4" runat="server" NavigateUrl="~/GetSample.aspx?Sample=4">Sample 4</asp:HyperLink>
        </td>
        <td>
            This sample demonstrates how to use a macro-enabled spreadsheet.
        </td>
        </tr>
        </table>
    </form>
</body>
</html>
