using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Drawing;
using System.Collections.Generic;
//using System.Web.UI.MobileControls;

public partial class WebUserControl_showPanel : System.Web.UI.UserControl
{
    List<VideoInfo> videoInfos = new List<VideoInfo>();
    private static SearchInfo searchInfo = new SearchInfo();
    public int totalCell = CommData.videoPerPage;
    public static int selectLength = 0;
    DBOperation dbSelect = new DBOperation();
    protected void Page_Load(object sender, EventArgs e)
    {        
        //     
        if (string.IsNullOrEmpty(searchInfo.Page))
        {
            searchInfo.Page = "1";
        }
    }
    public void AddSearchInfo( SearchInfo si)//first Load
    {
        searchInfo = si;
        videoInfos = dbSelect.SearchPatternSelectTable(searchInfo);
        selectLength = dbSelect.getVideoLength();
        ShowPage(1);
        ChangeSelectControl(0);
    }
    public void ShowVideoInfo(int page)
    {
        videoInfos = dbSelect.SearchPatternSelectTable(searchInfo);
        searchInfo.Page = page.ToString();
        /*foreach (VideoInfo i in videoInfos)
        {
            Response.Write(i.SessionSN + "<br>");
            Response.Write(i.File_FullName + "<br>");
            Response.Write(i.NameOfVideo + "<br>");
            foreach (string k in i.TopicList)
            {
                Response.Write(k + ' ');
            }
            Response.Write("<br><br>");
        }
        Response.Write("搜尋條件提示<br>VideoSessionType（RG精選；SS大會堂；GS我的學程；）：" + searchInfo.VideoSessionType + "<br>VideoLevel（0為不限等級）：" + searchInfo.VideoLevel + "<br>VideoInterestType（快速分類依序）：" + searchInfo.VideoInterestType + "<br>Page：" + searchInfo.Page + "<br>SearchPatthen（關鍵字搜尋）：" + searchInfo.SearchPatthen + "<br>RankType（rating評比最高；CourseCount看過最多；）：" + searchInfo.RankType + "<br>");        
        Response.Write("總筆數："+selectLength);*/
        if (videoInfos.Count == 0)
        {
            Response.Write("<script>alert('抱歉，查無相關條件錄影檔')</script>");
            HidePageControl();
            return;
        }
        else
        {
            List<WebUserControl_videoControl> tempVideoList = new List<WebUserControl_videoControl>();
            Label1.Text = "";
            div_has_show.Attributes.Add("style", "visibility:Hidden");
            foreach (VideoInfo videInfo in videoInfos)
            {

                WebUserControl_videoControl webUserControl_videoControl = (WebUserControl_videoControl)LoadControl("WebUserControl_videoControl.ascx");
                Label1.Text = " ■您已上过灰色区块之教材";
                div_has_show.Attributes.Add("style", "visibility:Visible");
                webUserControl_videoControl.SetVideoInfo(videInfo);
                tempVideoList.Add(webUserControl_videoControl);
            }
            int index =0;
            foreach (WebUserControl_videoControl webUserControl_videoControl in tempVideoList)
            {
                Table1.Rows[GetRowNumber(index)].Cells[GetColNumber(index)].Controls.Add(webUserControl_videoControl);
                index++;
            }
        }
        videoInfos.Clear();
    }
    private int GetRowNumber(int videoNumber)
    {
        return (videoNumber / 3);
    }
    private int GetColNumber(int videoNumber)
    {
        return (videoNumber % 3);
    }
    private void HidePageControl()
    {
        foreach (Control aControl in this.Controls)
        {
            if (aControl is LinkButton)
                aControl.Visible = false;
        }
    }
    public void ShowPage(int nowPage)
    {   
        HidePageControl();
        if (selectLength < 1)
            return;
        bool isZero = false;
        int totalPage = selectLength / totalCell;
        if ((selectLength % totalCell) > 0)
            totalPage++;
         int startpage = 0;
         if ((nowPage % 10) == 0)
         {
             isZero = true;
             startpage = nowPage - 9;
         }
         else
             startpage = nowPage - (nowPage % 10);
        int endpage = startpage + 10;
        if (endpage > totalPage)
            endpage = totalPage;
        for (int pageCount = startpage; pageCount < endpage; pageCount++)
        {
            ChangeControl(pageCount, endpage, true, isZero);
        }
        if (nowPage > 10)
        {
            LinkButton_firtPage.Visible = true;
            LinkButton_last10Page.Visible = true;
        }
      
        if ((totalPage - nowPage) > 10)
            LinkButton_next10Page.Visible = true;
  
        if (nowPage > 1)
            LinkButton_lastpage.Visible = true;
    
        if (nowPage < totalPage)
            LinkButton_nextPage.Visible = true;

    }

    private void ChangeControl(int pageCount,int maxpageCount,bool isVisible,bool isOnZero)
    {
        Control tempControl = FindControl("LinkButton_page" + ((pageCount % 10) + 1));
        if (tempControl != null)
        {
            LinkButton tempLinkButton = (LinkButton)tempControl;
            tempLinkButton.Visible = isVisible;
            if ((pageCount + 1) == maxpageCount && isOnZero)
                pageCount = pageCount - 10;
            tempLinkButton.Text = (pageCount + 1).ToString();
            tempLinkButton.ForeColor = Color.Black;
        }
    }
    private void ChangeSelectControl(int pageCount)
    {
        Control tempControl = FindControl("LinkButton_page" + ((pageCount % 10) + 1));
        if (tempControl != null)
        {
            LinkButton tempLinkButton = (LinkButton)tempControl;
            tempLinkButton.Visible = true;
            tempLinkButton.Text = "[ "+(pageCount + 1).ToString() +" ]";
            tempLinkButton.ForeColor = Color.Red;
            searchInfo.Page = (pageCount+1).ToString();
        }
    }
    public void SetPage(int page)
    {
        ShowVideoInfo(page+1);
        ChangeSelectControl(page);
    }
    protected void LinkButton_Click(object sender, EventArgs e)
    {
      
        int clickPage = 0;
        int oldPage =1;
        if (searchInfo != null)
        {
            oldPage = Convert.ToInt32(searchInfo.Page);
        }
        LinkButton tempLinkButton = (LinkButton)sender;
  
        if (tempLinkButton.Text.Equals("[第1頁]") || tempLinkButton.Text.Equals("[上10頁]") || tempLinkButton.Text.Equals("[下10頁]") || tempLinkButton.Text.Equals("[下一頁]") || tempLinkButton.Text.Equals("[上一頁]"))
        {
            if (tempLinkButton.Text.Equals("[第1頁]"))
            {
                clickPage = 1;
            }
            if (tempLinkButton.Text.Equals("[上10頁]"))
            {
                clickPage = oldPage - ((oldPage - 1) % 10) - 10;
            }
            if (tempLinkButton.Text.Equals("[下10頁]"))
            {
                clickPage = oldPage - ((oldPage - 1) % 10) + 10;
            }
            if (tempLinkButton.Text.Equals("[下一頁]"))
            {
                clickPage = oldPage + 1;
            }
            if (tempLinkButton.Text.Equals("[上一頁]"))
            {
                clickPage = oldPage -1;
            }

        }
        else
        {
            clickPage = Convert.ToInt32(tempLinkButton.Text.Replace("[", "").Replace("]", ""));
           
        }

        searchInfo.Page = clickPage.ToString();
        //Response.Write("AfterClickPage:"+clickPage + ' ' + searchInfo.Page);
        ShowPage( clickPage );
        SetPage( clickPage- 1 );
    }



}
