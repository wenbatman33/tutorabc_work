using System;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;

public partial class _Default : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

        //Session.Add("client_sn", 223543); 
        if (Session["client_sn"] == null)
            Response.Redirect("http://www.vipabc.com/program/member/member_login.asp");
        
        if (!IsPostBack)
        {
            SearchInfo searchInfo = new SearchInfo();
            CreateDefaultLevButton();
            //dbOperation.ClientSn = Session["client_sn"].ToString();
            //string strClientInterest = dbOperation.GetClientInterest();
            //string strClientLevel = dbOperation.GetClientLevel();
            //searchInfo.VideoInterestType = strClientInterest;
            //clientInterest.Value = strClientInterest;
            //searchInfo.VideoLevel = Convert.ToInt32(strClientLevel);
            //clientLevel.Value = strClientLevel;
            ShowtVideoInfo(searchInfo);
            ShowSearchInfo(searchInfo);

            string JavaScript = "JavaScript:";
            JavaScript+="var IE4 = (document.all) ? 1 : 0;";
            JavaScript += "window.document.getElementById('" + TextBox_videoNameInput.ClientID + "').value ='';";

            this.TextBox_videoNameInput.Attributes.Add("OnClick",JavaScript);
            this.TextBox_videoNameInput.Attributes.Add("onkeydown", "if(event.which || event.keyCode){if ((event.which == 13) || (event.keyCode == 13)) {document.getElementById('" + ImageButton_Search.UniqueID + "').click();return false;}} else {return true}; ");
        }
    }

    private void CreateDefaultLevButton()
    {
        for (int levCount = 1; levCount <= 12; levCount++)
        {
            CreateLevelButton(levCount);
        }
        
    }

 
    DBOperation dbOperation = new DBOperation();

    private string GetSerchPatten()
    {
        if (TextBox_videoNameInput.Text.Equals("搜尋字串可以顧問名稱，教材名稱來搜尋") || TextBox_videoNameInput.Text.Equals("搜寻字符串请以英文输入顾问名称或教材名称来搜寻"))
        {
            return "";
        }
        else
        {
           return TextBox_videoNameInput.Text;
        }
    }
   
    protected void LinkButton_GetSession_Click(object sender, EventArgs e)
    {
        SearchInfo searchInfo = new SearchInfo();
        LinkButton tempButton = (LinkButton)sender;
        string sessionType = "";
        if (tempButton.ID.Equals("LinkButton_mySession"))
            sessionType = "GS";
        if (tempButton.ID.Equals("LinkButton_regular"))
            sessionType = "RG";
        if (tempButton.ID.Equals("LinkButton_specialSession"))
            sessionType = "SS";
        if (tempButton.ID.Equals("LinkButton_allVideo"))
            sessionType = "";
        clientSessionType.Value = sessionType;
        searchInfo = CommData.setSearchInfo("1", sessionType, clientInterest.Value, clientLevel.Value, clientSearchPatten.Value, clientOrderBy.Value);
        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);

    }
    private void ShowSearchInfo(SearchInfo searchInfo)
    {
        ResetHttpLink();
        CreateDefaultLevButton();
        if (searchInfo != null)
        {
            foreach (Control aControl in form1.Controls)
            {
                if (aControl is LinkButton)
                {
                    LinkButton aTempLinkButton = (LinkButton)aControl;
                    if (searchInfo.VideoSessionType.Equals("GS")&&aTempLinkButton.ID.Equals("LinkButton_mySession"))
                        aTempLinkButton.ForeColor = Color.Red;
                    if (searchInfo.VideoSessionType.Equals("RG") && aTempLinkButton.ID.Equals("LinkButton_regular"))
                        aTempLinkButton.ForeColor = Color.Red;
                    if (searchInfo.VideoSessionType.Equals("SS") && aTempLinkButton.ID.Equals("LinkButton_specialSession"))
                        aTempLinkButton.ForeColor = Color.Red;
                    if (searchInfo.VideoSessionType.Equals("") && aTempLinkButton.ID.Equals("LinkButton_allVideo"))
                        aTempLinkButton.ForeColor = Color.Red;
                    if (searchInfo.RankType.Equals("rating") && aTempLinkButton.ID.Equals("LinkButton_highRating"))
                        aTempLinkButton.ForeColor = Color.Red;
                    if (searchInfo.RankType.Equals("CourseCount") && aTempLinkButton.ID.Equals("LinkButton_highView"))
                        aTempLinkButton.ForeColor = Color.Red;
                    /*if (aTempLinkButton.ID.Equals("LinkButton_" + searchInfo.Fous_Type))
                        aTempLinkButton.ForeColor = Color.Red;*/
                    if (searchInfo.VideoInterestType.Equals(clientInterest.Value) && aTempLinkButton.ID.Equals("LinkButton_topic_" + clientInterest.Value))
                        aTempLinkButton.ForeColor = Color.Red;
                    aTempLinkButton.CssClass = "HttpLinkStytle";
                }
                if (aControl is ImageButton)
                {
                     ImageButton aTempImageButton = (ImageButton)aControl;
                     //Response.Write("ImageButton_lev_" + clientLevel.Value + "<br>" + searchInfo.VideoLevel + "<br><br>");
                     if (aTempImageButton.ID.Equals("ImageButton_lev_" + clientLevel.Value) && searchInfo.VideoLevel.ToString().Equals(clientLevel.Value))
                     {
                         aTempImageButton.Attributes.Remove("onmouseover");
                         aTempImageButton.Attributes.Remove("onmouseout");
                         aTempImageButton.ImageUrl = "./Image/in_" + clientLevel.Value + ".gif";
                     }
                }
            }
           
        }
    }
    private void CreateLevelButton(int lev)
    {
        foreach (Control aControl in form1.Controls)
        {
            if (aControl is ImageButton)
            {
                ImageButton aTempImageButton = (ImageButton)aControl;
                if (aTempImageButton.ID.Equals("ImageButton_lev_" + lev))
                {
                    aTempImageButton.Attributes.Add("onmouseover", "this.src='./Image/in_" + lev + ".gif'");
                    aTempImageButton.Attributes.Add("onmouseout", "this.src='./Image/no_" + lev + ".gif'");
                    aTempImageButton.ImageUrl = "./Image/no_" + lev + ".gif";
                }
            }
        }
    }
    private void ResetHttpLink()
    {
        foreach (Control aControl in form1.Controls)
        {
            if(aControl is  LinkButton)
            {
                LinkButton aTempLinkButton = (LinkButton)aControl;
                aTempLinkButton.ForeColor = Color.Black;
            }
        }
    }
    private void ShowtVideoInfo(SearchInfo searchInfo)
    {
        dbOperation.ClientSn = Session["client_sn"].ToString();
        //因為原始定義的搜尋規則是下面，保留以做參考
        /*if (VideoSessionType.Equals("LinkButton_mySession"))
        {
            videoInfo = dbOperation.MySessionVideoList();
        }
        if (VideoSessionType.Equals("LinkButton_regular"))
        {
            videoInfo = dbOperation.AllRegSelectedVideoList();
        }
        if (VideoSessionType.Equals("LinkButton_specialSession"))
        {
            videoInfo = dbOperation.SelectedSpecialSessionVideoList();
        }
        if (VideoSessionType.Equals("LinkButton_allVideo"))
        {
            videoInfo = dbOperation.AllSelectedVideoList();
        }*/
        WebUserControl_showPanel1.AddSearchInfo(searchInfo);
        WebUserControl_showPanel1.ShowVideoInfo(1);
    }
    protected void LinkButtonTopic_Click(object sender, EventArgs e)
    {
        LinkButton tempButton = (LinkButton)sender;
        string[] data = tempButton.ID.Split('_');
        SearchInfo searchInfo = CommData.setSearchInfo("1", "", data[2], "", "", clientOrderBy.Value);             
        searchInfo.VideoInterestType = data[2];
        clientInterest.Value = data[2];
        //重置隱藏標籤
        clientSearchPatten.Value = "";
        clientLevel.Value = "";
        clientSessionType.Value = "";
        clientPage.Value = "1";

        TextBox_videoNameInput.Text = "搜寻字符串请以英文输入顾问名称或教材名称来搜寻";
        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);
        data = null;
    }
    protected void ImageButton_lev1_Click(object sender, ImageClickEventArgs e)
    {
        SearchInfo searchInfo = CommData.setSearchInfo("1", clientSessionType.Value, clientInterest.Value, clientLevel.Value, clientSearchPatten.Value, clientOrderBy.Value);
        ImageButton tempButton = (ImageButton)sender; 
        string[] data = tempButton.ID.Split('_');
        searchInfo.VideoLevel = Convert.ToInt32(data[2]);
        clientLevel.Value = data[2];
        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);
        data = null;

    }

    protected void ImageButton_Search_Click(object sender, ImageClickEventArgs e)
    {
        SearchInfo searchInfo = CommData.setSearchInfo("1", "", "", "", "", "");
        dbOperation.ClientSn = Session["client_sn"].ToString();
        if (TextBox_videoNameInput.Text.Equals("搜尋字串可以顧問名稱，教材名稱來搜尋") || TextBox_videoNameInput.Text.Equals("搜寻字符串请以英文输入顾问名称或教材名称来搜寻"))
        {
            searchInfo.SearchPatthen = "";
        }
        else
        {
            searchInfo.SearchPatthen = TextBox_videoNameInput.Text.Replace(" ","");//暫時先用去空白來搜尋0.0，如未來需求多關鍵字搜尋再改
        }
        //重置隱藏標籤
        clientSearchPatten.Value = searchInfo.SearchPatthen;
        clientInterest.Value = "";
        clientLevel.Value = "";
        clientSessionType.Value = "";
        clientPage.Value = "1";

        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);
    }
    protected void LinkButton_highRating_Click(object sender, CommandEventArgs e)
    {
        clientOrderBy.Value = e.CommandArgument.ToString();
        SearchInfo searchInfo = CommData.setSearchInfo("1", clientSessionType.Value, clientInterest.Value, clientLevel.Value, clientSearchPatten.Value, e.CommandArgument.ToString());
        /*if (!clientSearchPatten.Value.Equals(""))
        {
            searchInfo = CommData.setSearchInfo("1", "", "", "", clientSearchPatten.Value, e.CommandArgument.ToString());
        }*/
        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);
    }
    protected void LinkButton_highView_Click(object sender, CommandEventArgs e)
    {
        clientOrderBy.Value = e.CommandArgument.ToString();
        SearchInfo searchInfo = CommData.setSearchInfo("1", clientSessionType.Value, clientInterest.Value, clientLevel.Value, clientSearchPatten.Value, e.CommandArgument.ToString());
        /*if (!clientSearchPatten.Value.Equals(""))
        {
            searchInfo = CommData.setSearchInfo("1", "", "", "", clientSearchPatten.Value, e.CommandArgument.ToString());
        }*/
        ShowtVideoInfo(searchInfo);
        ShowSearchInfo(searchInfo);
    }
    //根據調查結果，這個select條件=沒作用，暫時不予理會
    protected void LinkButton_Focus(object sender, EventArgs e)
    {
        SearchInfo s = new SearchInfo();
        LinkButton tempButton = (LinkButton)sender;
        string[] data = tempButton.ID.Split('_');
        ShowtVideoInfo(s);
        data = null;
    }
}
