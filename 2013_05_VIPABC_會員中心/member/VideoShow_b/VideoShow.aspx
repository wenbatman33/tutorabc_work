<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="VideoShow.aspx.cs" Inherits="_Default" %>

<%@ Register src="UI/WebUserControl_videoControl.ascx" tagname="WebUserControl_videoControl" tagprefix="uc1" %>

<%@ Register src="UI/WebUserContro_RatingAndViewPanel.ascx" tagname="WebUserContro_RatingPanel" tagprefix="uc2" %>

<%@ Register src="UI/WebUserControl_showPanel.ascx" tagname="WebUserControl_showPanel" tagprefix="uc3" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<script type="text/javascript">
function mouseOver(imageIndex,controlId)
{
    document.getElementById(controlId).src ="lv_but"+imageIndex+"-on.gif";
}
function mouseOut(imageIndex,controlId)
{
    document.getElementById(controlId).src ="lv_but"+imageIndex+".gif";
}

</script>


    <title>更多精选咨询实况</title>
    <style type="text/css">
	a:hover{ 
	        text-decoration: underline;
        }
         .HttpLinkStytle
        {
          text-decoration: none;
        }
        #form1
        {
            height: 768px;
            width: 555px;
        }
        .style1
        {
            width: 555px;
            height: 749px;
            margin-right: 0px;
        }
        .style6
        {
            height: 20px;
        }
        .style8
        {
            height: 22px;
        }
        .style9
        {
            height: 28px;
            text-align: right;
        }
        .style10
        {            height: 58px;
                     background-image: url(./Image/searchbar.jpg);
        }
        .style13
        {
            font-size: small;
		
        }
        .style14
        {
            width: 73px;
            text-align: left;
           
        }
        .style15
        {
            /*height: 607px;*/
        }
    </style>
</head>
<body style="font-size: small">
<div>
    <form id="form1" runat="server" title="更多精选咨询实况">
    
    <table class="style1" >
        <tr>
            <td class="style10" colspan="2">
                <br />
                <br />
                <br />
            </td>
        </tr>
       <tr>
            <td  colspan="2">
			    <asp:Label ID="Label_warning_title" runat="server" Text="贴心提醒：" ForeColor="Red" ></asp:Label>
                <asp:Label ID="Label_warning_content" runat="server" Text="录像文件在线观看扣点后，于7天内重复在线观看相同录像文件时，不会被再次扣点！" ForeColor="#2B3F66"></asp:Label>

            </td>
        </tr>        
        <tr>
        
            <td class="style14" rowspan="5" valign="top" >
            <table style="width: 144%;">
                    <tr>
                        <td bgcolor="#FFFFFF">
                <span class="style13">
                <asp:Image ID="Image23" runat="server" Height="21px" 
                    ImageUrl="~/Image/search_03.jpg" Width="105px" />
				<br />
                <asp:LinkButton ID="LinkButton_topic_1" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">艺术娱乐</asp:LinkButton>
                <br />
        
                <asp:LinkButton ID="LinkButton_topic_2" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">生活趣闻</asp:LinkButton>
                <br />
               
                <asp:LinkButton ID="LinkButton_topic_3" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">运动休闲</asp:LinkButton>
                <br />
          
                <asp:LinkButton ID="LinkButton_topic_4" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">教育训练</asp:LinkButton>
                <br />
         
                <asp:LinkButton ID="LinkButton_topic_5" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">科学计算机</asp:LinkButton>
                <br />
         
                <asp:LinkButton ID="LinkButton_topic_6" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">世界地理</asp:LinkButton>
                <br />
         
                <asp:LinkButton ID="LinkButton_topic_7" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">观光旅游</asp:LinkButton>
                <br />
            
                <asp:LinkButton ID="LinkButton_topic_8" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">生活用语</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_9" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">商业金融</asp:LinkButton>
                <br />
         
                <asp:LinkButton ID="LinkButton_topic_10" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">管理相关</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_11" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">健康医疗</asp:LinkButton>
                <br />
        
                <asp:LinkButton ID="LinkButton_topic_12" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">流行时尚</asp:LinkButton>
                <br />
     
                <asp:LinkButton ID="LinkButton_topic_13" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">人际关系</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_14" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">政治法律</asp:LinkButton>
                <br />
  
                <asp:LinkButton ID="LinkButton_topic_15" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">居家生活</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_16" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">工作职涯</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_17" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">城市国家</asp:LinkButton>
                <br />
     
                <asp:LinkButton ID="LinkButton_topic_18" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">历史文化</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_19" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">动物植物</asp:LinkButton>
                <br />

                <asp:LinkButton ID="LinkButton_topic_20" runat="server" 
                    onclick="LinkButtonTopic_Click" ForeColor="Black">美食烹饪</asp:LinkButton>
                <br />
                            <br />
                </span>
                        </td>
                       
                    </tr>
                    <tr>
                        <td>&nbsp;
                            </td>
                       
                    </tr>
                   
                </table>
                <span>
                <br />
                </span>
                <br />
                
            </td>
            <td bgcolor="#F8FBFC">
                <asp:TextBox ID="TextBox_videoNameInput" runat="server" Height="21px"
                    Width="318px">搜寻字符串请以英文输入顾问名称或教材名称来搜寻</asp:TextBox>
                <asp:ImageButton ID="ImageButton_Search" runat="server" 
                    ImageUrl="~/Image/search-btn_09.gif" onclick="ImageButton_Search_Click"  />
                </td>
        </tr>
        <tr>
            <td class="style6" bgcolor="#F3F9F9" >
&nbsp;<asp:LinkButton ID="LinkButton_regular" runat="server" 
                    onclick="LinkButton_GetSession_Click">精选</asp:LinkButton>
                &nbsp;&nbsp;|
                <asp:LinkButton ID="LinkButton_specialSession" runat="server" 
                    onclick="LinkButton_GetSession_Click">大会堂</asp:LinkButton>
&nbsp;|
                <asp:LinkButton ID="LinkButton_mySession" runat="server" 
                    onclick="LinkButton_GetSession_Click">我的学程</asp:LinkButton>
&nbsp;|
                <asp:LinkButton ID="LinkButton_allVideo" runat="server" 
                    onclick="LinkButton_GetSession_Click">全部</asp:LinkButton>
            </td>
        </tr>
        <tr>
            <td class="style8" align=left bgcolor="#F8FBFC">
                <asp:Label ID="Label1" runat="server" Text="Level:"></asp:Label>
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_1" runat="server" 
                    ImageUrl="~/Image/lv_but1.gif" onclick="ImageButton_lev1_Click" 
                     />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_2" runat="server" 
                    ImageUrl="~/Image/lv_but2.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_3" runat="server" 
                    ImageUrl="~/Image/lv_but3.gif" onclick="ImageButton_lev1_Click" 
                     />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_4" runat="server" 
                    ImageUrl="~/Image/lv_but4.gif" onclick="ImageButton_lev1_Click" 
                    />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_5" runat="server" 
                    ImageUrl="~/Image/lv_but5.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_6" runat="server" 
                    ImageUrl="~/Image/lv_but6.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_7" runat="server" 
                    ImageUrl="~/Image/lv_but7.gif" onclick="ImageButton_lev1_Click" 
                    />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_8" runat="server" 
                    ImageUrl="~/Image/lv_but8.gif" onclick="ImageButton_lev1_Click" 
                     />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_9" runat="server" 
                    ImageUrl="~/Image/lv_but9.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_10" runat="server" 
                    ImageUrl="~/Image/lv_but10.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_11" runat="server" 
                    ImageUrl="~/Image/lv_but11.gif" onclick="ImageButton_lev1_Click" />
                    &nbsp;<asp:ImageButton ID="ImageButton_lev_12" runat="server" 
                    ImageUrl="~/Image/lv_but12.gif" onclick="ImageButton_lev1_Click" 
                   />
            </td>
        </tr>
        <tr>
            <td class="style9" align="left" bgcolor="#FBFCFE">
                &nbsp;<!--<asp:LinkButton ID="LinkButton_Grammar" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus" >文法</asp:LinkButton>
&nbsp;<asp:LinkButton ID="LinkButton_Vocabulary" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus">字汇</asp:LinkButton>
&nbsp;<asp:LinkButton ID="LinkButton_Pronuciation" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus">发音</asp:LinkButton>
&nbsp;<asp:LinkButton ID="LinkButton_Speaking" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus">会话</asp:LinkButton>
&nbsp;<asp:LinkButton ID="LinkButton_Listening" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus">听力</asp:LinkButton>
&nbsp;<asp:LinkButton ID="LinkButton_Reading" runat="server" ForeColor="Black" 
                    onclick="LinkButton_Focus">阅读</asp:LinkButton>-->
                 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; 
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:LinkButton ID="LinkButton_highRating" runat="server" ForeColor="Black" 
                   CommandArgument="rating" OnCommand="LinkButton_highRating_Click" >评比最高</asp:LinkButton>
                |
                <asp:LinkButton ID="LinkButton_highView" runat="server" ForeColor="Black" 
                    CommandArgument="CourseCount" OnCommand="LinkButton_highView_Click">看过最多</asp:LinkButton>
            </td>
        </tr>
        <tr>
            <td class="style15">
                <uc3:WebUserControl_showPanel ID="WebUserControl_showPanel1" runat="server" />
            </td>
        </tr>
    </table>
    <asp:hiddenfield  id="clientInterest" value="" runat="server" ></asp:hiddenfield>
    <asp:hiddenfield  id="clientSessionType" value="" runat="server" ></asp:hiddenfield>
    <asp:hiddenfield  id="clientSearchPatten" value="" runat="server" ></asp:hiddenfield>
    <asp:hiddenfield  id="clientLevel" value="" runat="server" ></asp:hiddenfield>
    <asp:hiddenfield  id="clientOrderBy" value="" runat="server" ></asp:hiddenfield>
    <asp:hiddenfield  id="clientPage" value="" runat="server" ></asp:hiddenfield>
    
    </form>
    </div>

</body>
</html>
