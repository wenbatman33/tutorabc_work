<%@ Control Language="C#" AutoEventWireup="true" CodeFile="WebUserContro_RatingAndViewPanel.ascx.cs" Inherits="WebUserContro_RatingPanel" %>
<link rel="stylesheet" href="/css/css_cn.css" type="text/css">
<style type="text/css">
    .msg  {float:right; width:160px;font-size:11px; line-height:18px;  height:22px;  padding:0; margin:0; }
    .style1
    {
        width: 311px;
    }
    .style2
    {
        width: 170px;
    }
    .style3
    {
        width: 177px;
    }
    .style4
    {
        width: 190px;
    }
</style>

<asp:Panel ID="Panel1" runat="server" Height="297px" Width="304px">
    <div class="pop-reco">
    <div class="msg">
        <asp:Label ID="Label2" runat="server" Text="Label" Font-Size="12px"></asp:Label>                
    </div>
	<div class="photo">
    <div class="btn"><asp:Button ID="Button_reviewThisSession" runat="server" Text="扣点回顾此课程"   Visible="False" onclick="Button_reviewThisSession_Click" class="btn_1" />
	</div> 	
	<asp:Image ID="Image1" runat="server" Height="80px" Width="120px" />
	</div>
	<div class="wo">
                <asp:Label ID="Label_nameOFVideo" runat="server" 
                    style="text-decoration: underline" Text="课程名称"></asp:Label>
    
                <br />
    
                <asp:Label ID="Label_consultant0" runat="server" Text="顾问"></asp:Label>
                :<asp:Label ID="Label_consultant" runat="server" 
                    style="text-decoration: underline" Text="Consultant"></asp:Label>
         
                <br />
         
                <asp:Label ID="Label_lv" runat="server" Text="等级:"></asp:Label>
                <asp:Image ID="Image_lv" runat="server" Height="9px" Width="10px" />
                &nbsp;<br /> 评比:<asp:Label ID="Label_rating" runat="server" ForeColor="Red" Text="8.5"></asp:Label>
             
                <br />
             
                <asp:Label ID="Label1" runat="server" Text="使用堂数 : "></asp:Label>
                <asp:Label ID="Label_sessionCost" runat="server" Text="0.25"></asp:Label>
     
                <br />
     
                <asp:DropDownList ID="DropDownList_rating" runat="server" Height="16px" 
                    Visible="False" Width="48px">
                    <asp:ListItem>1</asp:ListItem>
                    <asp:ListItem>2</asp:ListItem>
                    <asp:ListItem>3</asp:ListItem>
                    <asp:ListItem>4</asp:ListItem>
                    <asp:ListItem>5</asp:ListItem>
                    <asp:ListItem Value="6"></asp:ListItem>
                    <asp:ListItem>7</asp:ListItem>
                    <asp:ListItem>8</asp:ListItem>
                    <asp:ListItem>9</asp:ListItem>
                    <asp:ListItem>10</asp:ListItem>
                </asp:DropDownList>
                &nbsp;<asp:Button ID="Button_rating" runat="server" onclick="Button_rating_Click" 
                    Text="评比" Visible="False" />
       
                <br />
       
                课后评鉴分数：<asp:Label ID="Label_sessionRating" runat="server" Text="0"></asp:Label>
   


                <br />
   


    <asp:ListBox ID="ListBox_clientRating" runat="server" Height="30px" 
        Visible="False" Width="30px"></asp:ListBox>
	</div>
	</div>
</asp:Panel>



