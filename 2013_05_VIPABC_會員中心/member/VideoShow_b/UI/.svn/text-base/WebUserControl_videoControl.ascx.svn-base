<%@ Control Language="C#" AutoEventWireup="true" CodeFile="WebUserControl_videoControl.ascx.cs" Inherits="WebUserControl_videoControl" %>

<%@ Register src="WebUserControl_ShowRateing.ascx" tagname="WebUserControl_ShowRateing" tagprefix="uc1" %>


<asp:Panel ID="Panel_videoShow" runat="server" Height="198px" Width="149px" 
    BorderWidth="0px" Font-Size="Small" HorizontalAlign="Left" 
    BackColor="#F7FBFF" BorderStyle="None">
        <style type="text/css">
        .HttpLinkStytle
        {
          text-decoration: none;
        }
        </style>
        <asp:Image ID="Image_thumb" runat="server" Height="85px" ImageAlign="Top" 
            Width="142px" BorderColor="#CCCCCC" BorderStyle="Inset" 
            BorderWidth="2px" />
        <br />
        <div style="word-wrap: break-word;word-break: break-all;">
        <asp:HyperLink ID="HyperLink_nameOfvideo" runat="server" Font-Bold="True" 
            Height="37px" Width="139px" EnableTheming="False" Font-Names="Arial" 
            Font-Size="9pt" CssClass="HttpLinkStytle">HyperLink</asp:HyperLink>
            </div>
        <asp:Label ID="Label_consultant" runat="server" Text="顾问：" Font-Size="Small"></asp:Label>
    <asp:Label ID="Label_consultantName" runat="server" Text="Yahoo" Font-Size="Small" 
            Width="95px"></asp:Label>
    <br />
        <asp:Label ID="Label_totalViewCount" runat="server" Text="观看次数：" 
            Font-Size="Small"></asp:Label>
    <asp:Label ID="Label_viewCount" runat="server" Text="1,000" Font-Size="Small"></asp:Label>
    <br />
        <asp:Label ID="Label5" runat="server" Text="使用堂数：" Font-Size="Small"></asp:Label>
    <asp:Label ID="Label_sessionCost" runat="server" Text="0.5" Font-Size="Small"></asp:Label>
        <uc1:WebUserControl_ShowRateing ID="WebUserControl_ShowRateing1" 
            runat="server" />
</asp:Panel>
