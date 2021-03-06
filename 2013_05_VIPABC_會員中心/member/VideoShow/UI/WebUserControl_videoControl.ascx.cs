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
using System.IO;
using System.Drawing;

public partial class WebUserControl_videoControl : System.Web.UI.UserControl
{

    private VideoInfo readVideoInfo = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        this.Image_thumb.Attributes.Add("onmouseover", "document.body.style.cursor = 'hand';");
        this.Image_thumb.Attributes.Add("onmouseout", "document.body.style.cursor = 'default';");
    }

    public void SetVideoInfo(VideoInfo videoInfo)
    {

        Base64 base64 = new Base64();
        readVideoInfo = videoInfo;
        string csn = "";
        string ratingAction = CryptLib.Crypt_Tool.EncryptAnyKeyStr("ratingAction", "ctabcpc07");
        string viewAction = CryptLib.Crypt_Tool.EncryptAnyKeyStr("viewAction", "ctabcpc07");
        string sessionSn = CryptLib.Crypt_Tool.EncryptAnyKeyStr(videoInfo.SessionSN, "ctabcpc07");

        if (Session["client_sn"] != null)
        {
            csn = CryptLib.Crypt_Tool.EncryptAnyKeyStr(Session["client_sn"].ToString(), "ctabcpc07");
            Label_consultantName.Text = CreatePaddingString(videoInfo.Consultant, 120);
            Label_consultantName.ToolTip = videoInfo.Consultant;
            Image_thumb.ToolTip = videoInfo.NameOfVideo;
            HyperLink_nameOfvideo.ToolTip = videoInfo.NameOfVideo;
            //HyperLink_rating.Text= videoInfo.Mtl_Rating.ToString("#.##");
            //HyperLink_rating.Attributes.Add("onclick", "window.open('VideoRating.aspx?sn=" + base64.EncodeTo64(sessionSn) + "&csn=" + base64.EncodeTo64(csn) + "&action=" + base64.EncodeTo64(ratingAction) + "', '諮詢實況', 'height=380, width=320, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=yes,location=no, status=no')");
            //HyperLink_rating.Attributes.Add("href", "javascript:;");
            this.HyperLink_nameOfvideo.Attributes.Add("onclick", "newwindow=window.open('VideoRating.aspx?IsClientUseThisSessionSn=" + videoInfo.IsClientUseThisSessionSn + "&IsClientUseThisMaterial=" + videoInfo.IsClientUseThisMaterial + "&IsClientUseThisRecordFile=" + videoInfo.IsClientUseThisRecordFile + "&sn=" + base64.EncodeTo64(sessionSn) + "&csn=" + base64.EncodeTo64(csn) + "&action=" + base64.EncodeTo64(viewAction) + "', '諮詢實況', 'height=320, width=400 top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=yes,location=no, status=no');if (window.focus) {newwindow.focus()}");
            HyperLink_nameOfvideo.Attributes.Add("href", "javascript:void(0);");
            Image_thumb.Attributes.Add("onclick", "newwindow=window.open('VideoRating.aspx?IsClientUseThisSessionSn=" + videoInfo.IsClientUseThisSessionSn + "&IsClientUseThisMaterial=" + videoInfo.IsClientUseThisMaterial + "&IsClientUseThisRecordFile=" + videoInfo.IsClientUseThisRecordFile + "&sn=" + base64.EncodeTo64(sessionSn) + "&csn=" + base64.EncodeTo64(csn) + "&action=" + base64.EncodeTo64(viewAction) + "', '諮詢實況', 'height=320, width=400, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=yes,location=no, status=no');if (window.focus) {newwindow.focus()}");
            Image_thumb.Attributes.Add("href", "javascript:void(0);");
            this.HyperLink_nameOfvideo.Text = CreatePaddingString(videoInfo.NameOfVideo, 400);
            Label_viewCount.Text = videoInfo.CourseCount.ToString();
        }

        //====== 工單:BD09050122、BD09050603 特殊扣點觀看錄影檔設定 raychien 20090511 START ======
        //' GetPointsFromDownloadVideoRule
        //'             回傳 -2，表示沒有為此客戶設定任何扣點規則
        //'             回傳 -3，表示傳入的ClientSn或VideoType是空值
        //'             回傳 -4，表示為此客戶設定的規則的下載次數已達上限
        DBOperation dbOperation = new DBOperation();
        float vds_subtract_points = 0;
        if (videoInfo.SessionType.Equals("SS") && !videoInfo.IsOwnSession)
        {
            Label_sessionCost.Text = "1.0 堂";
            vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 3, "", "");
            if (vds_subtract_points != -2 && vds_subtract_points != -4)
            {
                if ( vds_subtract_points >= 0 )
                {
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                }
            }
        }
        else
        {
            if (videoInfo.IsOwnSession)
            {
                Label_sessionCost.Text = "0 堂";
                vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 1, "", "");
            }
            else
            {
                Label_sessionCost.Text = "1.0 堂";
                vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 2, "", "");
            }

            if ( vds_subtract_points != -2 && vds_subtract_points != -4 )
            {
                if ( vds_subtract_points >= 0 )
                {
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                }
            }
        }
        //====== 工單:BD09050122、BD09050603 特殊扣點觀看錄影檔設定 raychien 20090511 END ======
        if (videoInfo.LVList.Count > 0 && videoInfo.SessionLevel == 0)
        {
            string lv = videoInfo.LVList[0].ToString();
            string lv_image_Path = "../Image/lv" + lv + ".gif";
            //Image_lv.ImageUrl = lv_image_Path;
        }
        else if (videoInfo.SessionLevel != 0)
        {
            string lv_image_Path = "../Image/lv" + videoInfo.SessionLevel + ".gif";
            //Image_lv.ImageUrl = lv_image_Path;
        }
        if (!File.Exists(ConfigurationManager.AppSettings["mtl_path"].ToString() + Path.DirectorySeparatorChar + videoInfo.Course + ".jpg"))
        {
            this.Image_thumb.ImageUrl = "../Image/mov.jpg";
        }
        else
        {
            this.Image_thumb.ImageUrl = "/images/mtl_thumb/" + videoInfo.Course + ".jpg";
        }

        //已上過此堂課               
        if (videoInfo.IsClientUseThisSessionSn && !videoInfo.IsClientUseThisRecordFile) 
        {
            this.Panel_videoShow.BackColor = Color.Gainsboro;
        }
       //曾經觀看過
        else if (videoInfo.IsClientUseThisSessionSn && videoInfo.IsClientUseThisRecordFile) 
        {
            this.Panel_videoShow.BackColor = Color.Gainsboro;
        }
       //已上過此堂課 (相同教材)
        else if (!videoInfo.IsClientUseThisSessionSn && videoInfo.IsClientUseThisMaterial && !videoInfo.IsClientUseThisRecordFile) 
        {
            this.Panel_videoShow.BackColor = Color.Gainsboro;
        }
       //曾經觀看過
        else if (!videoInfo.IsClientUseThisSessionSn && videoInfo.IsClientUseThisMaterial && videoInfo.IsClientUseThisRecordFile) 
        {
            this.Panel_videoShow.BackColor = Color.Gainsboro;
        }
        //曾經觀看過
        else if (!videoInfo.IsClientUseThisSessionSn && !videoInfo.IsClientUseThisMaterial && videoInfo.IsClientUseThisRecordFile) 
        {
            this.Panel_videoShow.BackColor = Color.Gainsboro;
        }
        WebUserControl_ShowRateing1.SetRating((float)videoInfo.Mtl_Rating);
    }

    private string CreatePaddingString(string orgText, float maxLen)
    {
        float totalLen = 10.8671875f;
        char[] chars = orgText.ToCharArray();
        int tesxCount = 0;
        for (int stringIndex = 0; stringIndex < chars.Length; stringIndex++)
        {
            totalLen += GetTextSize(chars[stringIndex].ToString());
            if (totalLen > maxLen)
                break;
            tesxCount++;
        }
        if (tesxCount == orgText.Length)
            return orgText;
        return orgText.Substring(0, tesxCount - 1) + "..";
    }

    private float GetTextSize(string aChar)
    {
        Bitmap bmp = new Bitmap(139, 37);
        Graphics g = Graphics.FromImage(bmp);
        SizeF size = g.MeasureString(aChar, new Font("Arial", 9, FontStyle.Bold));
        bmp.Dispose();
        g.Dispose();
        return size.Width;
    }

    public WebUserControl_videoControl()
    {
    }
}
