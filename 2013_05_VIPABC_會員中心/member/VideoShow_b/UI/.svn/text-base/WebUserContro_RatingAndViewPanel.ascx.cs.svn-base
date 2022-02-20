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
using System.Net;

public partial class WebUserContro_RatingPanel : System.Web.UI.UserControl
{
    DBOperation dbOperation = new DBOperation();

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["client_sn"] == null)
            Response.Redirect("VideoShowForm.asp?language=zh-tw");
        if (Request["csn"] == null)
        {
            Response.Redirect("VideoShowForm.asp?language=zh-tw");
        }
        Base64 base64 = new Base64();
        string csn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["csn"].ToString()), "ctabcpc07");
        if (!csn.Equals(Session["client_sn"].ToString()))
            Response.Redirect("VideoShowForm.asp?language=zh-tw");
    }

    public void SetVideoInfo(VideoInfo videoInfo, string clientSN, bool isShowViewButton)
    {
        if (!IsRecodeVideoExist(videoInfo))
        {
            Response.Write("<Script>alert('抱歉，系统维修中，请稍后再试，谢谢');window.opener = self; self.close();</Script>");
            return;
        }

        double usedPoint = 0;
        dbOperation.ClientSn = clientSN;
        Label_consultant.Text = videoInfo.Consultant;
        Label_nameOFVideo.Text = videoInfo.NameOfVideo;
        Label_rating.Text = videoInfo.Mtl_Rating.ToString("#.##");//不明的評比邏輯
        Label_sessionRating.Text = videoInfo.SessionRating.ToString();
        Label2.Text = "";

        //====== 工单:BD09050122、BD09050603 特殊扣点观看录像文件设定 raychien 20090511 START ======
        float vds_subtract_points = 0, vds_subtract_points_all = 0 ;
        //大会堂录像文件
        if (videoInfo.SessionType.Equals("SS") && !videoInfo.IsOwnSession)
        {
            Label_sessionCost.Text = "1.0";
            usedPoint = 1;
            vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 3, "", "");
            if (vds_subtract_points != -2 && vds_subtract_points != -4)
            {
                if (vds_subtract_points >= 0)
                {
                    //Label_sessionCost.Text = "1.0";
                    //usedPoint = 1;
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                    usedPoint = (float)Convert.ToInt32(vds_subtract_points / 65);
                }
                else
                {
                    Response.Write("<Script>alert('录像文件点数有误，请洽客服');window.opener = self; self.close();</Script>");
                    return;
                }
            }
        }
        else
        {
            // 个人录像文件
            if (videoInfo.IsOwnSession)
            {
                usedPoint = 0;
                Label_sessionCost.Text = "0";
                vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 1, "", "");
                if ( vds_subtract_points >= 0 )
                {
                    //Label_sessionCost.Text = "1.0";
                    //usedPoint = 1;
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                    usedPoint = (float)Convert.ToInt32(vds_subtract_points / 65);
                }
            }
            // 精选录像文件
            else
            {
                usedPoint = 1.0;
                Label_sessionCost.Text = "1.0";
                vds_subtract_points = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 2, "", "");
                if ( vds_subtract_points >= 0 )
                {
                    //Label_sessionCost.Text = "1.0";
                    //usedPoint = 1;
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                    usedPoint = (float)Convert.ToInt32(vds_subtract_points / 65);
                }
            }
            if (vds_subtract_points != -2 && vds_subtract_points != -4)
            {
                if ( vds_subtract_points >= 0 )
                {
                    //Label_sessionCost.Text = "1.0";
                    //usedPoint = 1;
                    Label_sessionCost.Text = Convert.ToString(vds_subtract_points / 65);
                    usedPoint = (float)Convert.ToInt32(vds_subtract_points / 65);
                }
                else
                {
                    Response.Write("<Script>alert('录像文件点数有误，请洽客服');window.opener = self; self.close();</Script>");
                    return;
                }
            }
        }
        //====== 工单:BD09050122、BD09050603 特殊扣点观看录像文件设定 raychien 20090511 END ======

        //'20130207 阿捨新增 不拘類型上限判斷
        vds_subtract_points_all = dbOperation.GetPointsFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), 9, "", "");

        if (videoInfo.Mtl_RatingInfo.IsView && videoInfo.Mtl_RatingInfo.UserRating < 0)
        {
            DropDownList_rating.Visible = true;
            Button_rating.Visible = true;
        }
        
        if (videoInfo.LVList.Count > 0)
        {
            string lv="0";
            foreach (String vi in videoInfo.LVList)
            {
               if(!string.IsNullOrEmpty(vi))//因為此view加了前後逗號，在存到ArrayList時會存兩個空物件
                   lv = vi.ToString();
            }
            string lv_image_Path = "../Image/lv" + lv + ".gif";
            Image_lv.ImageUrl = lv_image_Path;
        }
        else if (videoInfo.SessionLevel != 0)
        {
            string lv_image_Path = "../Image/lv" + videoInfo.SessionLevel + ".gif";
            Image_lv.ImageUrl = lv_image_Path;
        }

        //====== 工单:RD11062102 wording
        if (Request["IsClientUseThisSessionSn"] == null)
        {
            return;
        }
        if (Request["IsClientUseThisMaterial"] == null)
        {
            return;
        }
        if (Request["IsClientUseThisRecordFile"] == null)
        {
            return;
        }

        //已上过此堂课               
        if (Request["IsClientUseThisSessionSn"].Equals("True") && Request["IsClientUseThisRecordFile"].Equals("False"))
        {
            Label2.Text = "★已上过此堂课";
        }
        //曾经观看过
        else if (Request["IsClientUseThisSessionSn"].Equals("True") && Request["IsClientUseThisRecordFile"].Equals("True"))
        {
            Label2.Text = "★曾经观看过";
        }
        //已上过此堂课 (相同教材)
        else if (Request["IsClientUseThisSessionSn"].Equals("False") && Request["IsClientUseThisMaterial"].Equals("True") && Request["IsClientUseThisRecordFile"].Equals("False"))
        {
            Label2.Text = "★已上过此堂课（相同教材）";
        }
        //曾经观看过
        else if (Request["IsClientUseThisSessionSn"].Equals("False") && Request["IsClientUseThisMaterial"].Equals("True") && Request["IsClientUseThisRecordFile"].Equals("True"))
        {
            Label2.Text = "★曾经观看过";
        }
        //曾经观看过
        else if (Request["IsClientUseThisSessionSn"].Equals("False") && Request["IsClientUseThisMaterial"].Equals("False") && Request["IsClientUseThisRecordFile"].Equals("True"))
        {
            Label2.Text = "★曾经观看过";
        }
        if (!File.Exists(ConfigurationManager.AppSettings["mtl_path"].ToString() + Path.DirectorySeparatorChar + videoInfo.Course + ".jpg"))
        {
            this.Image1.ImageUrl = "../Image/mov.jpg";
        }
        else
        {
            this.Image1.ImageUrl = "/images/mtl_thumb/" + videoInfo.Course + ".jpg";
        }

        if (isShowViewButton)
            Button_reviewThisSession.Visible = true;

        double freePoint = dbOperation.GetUserFreePoint(clientSN);

        if (!dbOperation.IsContractValid())
        {
            Button_reviewThisSession.Text = "合约已到期，无法观看";
            Button_reviewThisSession.Enabled = false;
            return;
        }
        if (dbOperation.IsUnlimitDownloadTime(videoInfo, clientSN) || dbOperation.IsFreeTimeDonlwoad())
        {
            Button_reviewThisSession.Text = "免费复习此课程";
        }
        else if (!dbOperation.IsPointEnough(usedPoint) && freePoint < usedPoint)
        {
            Button_reviewThisSession.Text = "点数不足，抱歉";
            Button_reviewThisSession.Enabled = false;
        }
		else if (freePoint >= usedPoint && vds_subtract_points <= 0 )
        {
            Button_reviewThisSession.Text = "从赠送堂数中扣点";
        }
        else if ( vds_subtract_points == -4 )
        {
            Button_reviewThisSession.Text = "点数不足，抱歉";
            Button_reviewThisSession.Enabled = false;
        }
        else if (vds_subtract_points_all == -4 || vds_subtract_points_all == -5 )
        {
            Button_reviewThisSession.Text = "点数不足，抱歉";
            Button_reviewThisSession.Enabled = false;
        }
        /*ArrayList userRatingList = new ArrayList();
        Response.Write(videoInfo.SessionSN + ' ' + videoInfo.Course + ' ' + videoInfo.Mtl_RatingInfo.UsersRatingDate);
        //根本沒做...
        if (videoInfo.Mtl_RatingInfo.UsersRatingDate.Count > 0)
        {
            ListBox_clientRating.Visible = true;
            for (int ratingIndex = 0; ratingIndex < videoInfo.Mtl_RatingInfo.UsersRatingDate.Count; ratingIndex++)
            {
                userRatingList.Add(videoInfo.Mtl_RatingInfo.UsersRatingDate[ratingIndex] + "," + videoInfo.Mtl_RatingInfo.UsersRatingList[ratingIndex]);
            }
            userRatingList.Sort();
            foreach (string aRatingString in userRatingList)
            {
                string[] ratringDatas = aRatingString.Split(',');
                DateTime rateingDate = DateTime.Parse(ratringDatas[0]);
                ListBox_clientRating.Items.Add(rateingDate.ToShortDateString() + "                 " + ratringDatas[1]);
            }
        }*/
    }

    protected void Button_rating_Click(object sender, EventArgs e)
    {
        Base64 base64 = new Base64();
        if (Request["sn"] == null)
        {
            return;
        }
        if (Request["csn"] == null)
        {
            return;
        }
        string sn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["sn"].ToString()), "ctabcpc07");
        string csn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["csn"].ToString()), "ctabcpc07");
        dbOperation.ClientSn = csn;
        VideoInfo videoInfo = dbOperation.GetVideoInfoBySessionSN(sn);
        dbOperation.AddUserRating(videoInfo, Convert.ToInt32(DropDownList_rating.SelectedItem.Value));
        DropDownList_rating.Visible = false;
        Button_rating.Visible = false;
    }

    protected void Button_reviewThisSession_Click(object sender, EventArgs e)
    {
        Base64 base64 = new Base64();
        string sn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["sn"].ToString()), "ctabcpc07");
        string csn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["csn"].ToString()), "ctabcpc07");
        dbOperation.ClientSn = csn;
        VideoInfo videoInfo = dbOperation.GetVideoInfoBySessionSN(sn);

        if (!IsRecodeVideoExist(videoInfo))
        {
            Response.Write("<Script>alert('抱歉，系统维修中，请稍后再试');window.opener = self; self.close();</Script>");
            return;
        }

        double freePoint = dbOperation.GetUserFreePoint(csn);
        double usedPoint = Convert.ToDouble(Label_sessionCost.Text);
        bool is_in_7Days = dbOperation.IsUnlimitDownloadTime(videoInfo, csn);

        if (freePoint >= usedPoint && usedPoint != 0)
        {
            if (is_in_7Days)
            {
                dbOperation.AddtoFreeViewList(videoInfo, csn, usedPoint * 65.0, Request.UserHostAddress);
            }
            // 扣的是赠送点数
            else
            {
                dbOperation.AddtoFreePointViewList(videoInfo, csn, usedPoint * 65.0, Request.UserHostAddress);
                dbOperation.DecreaseFreePoint(videoInfo, csn, (float)usedPoint);
            }
        }
        else
        {
            if (is_in_7Days || dbOperation.IsFreeTimeDonlwoad())
            {
                dbOperation.AddtoFreeViewList(videoInfo, csn, usedPoint * 65.0, Request.UserHostAddress);
            }
            // 扣的是一般的点数
            else
            {
                //====== 工单:BD09050122、BD09050603 特殊扣点观看录像文件设定 raychien 20090511 START ======

                float vds_subtract_points = 0;
                string vds_vdieo_type = "";
                string vds_seq = "";
               
                // 大会堂录像文件
                if (videoInfo.SessionType.Equals("SS") && !videoInfo.IsOwnSession)
                {
                    vds_vdieo_type = "3";
                }
                else
                {
                    // 个人录像文件
                    if (videoInfo.IsOwnSession)
                    {
                        vds_vdieo_type = "1";
                    }
                    else
                    {
                        vds_vdieo_type = "2";
                    }
                }
                vds_seq = dbOperation.GetSeqFromDownloadVideoRule(Convert.ToInt32(Session["client_sn"]), vds_vdieo_type);
                if (Convert.ToInt32(vds_seq) < 0)
                    vds_seq = "";
                dbOperation.AddtoCostViewList(videoInfo, csn, usedPoint * 65.0, Request.UserHostAddress, vds_vdieo_type, vds_seq);
                
                //====== 工单:BD09050122、BD09050603 特殊扣点观看录像文件设定 raychien 20090511 END ======
            }
        }
        DropDownList_rating.Visible = true;
        Button_rating.Visible = true;
        Button_reviewThisSession.Text = "免费复习此课程";

        //20091029 pipi TutorMeet ======= start ========
        //=========播放的地方==========
        int tutormeetShow = videoInfo.File_FullName.IndexOf(".jnr");

        if (tutormeetShow > 0)
        {
            // Response.Redirect("TutorMeetShowVideo.aspx?videoname=_recording_session903_88733_2009061719903");
            Response.Redirect(GetVideoFileUrl(videoInfo), false);
        }
        else
        {
            Response.Redirect("TutorMeetShowVideo.aspx?videoname=" + videoInfo.File_FullName + "");
        }
        //20091029 pipi TutorMeet ======= end ========
    }

    private string GetVideoFileUrl(VideoInfo videoInfo)
    {
        return ConfigurationManager.AppSettings["videoRecodeServer"].ToString() + "/svr" + videoInfo.File_Path + "/" + videoInfo.File_Date.Year + "_" + videoInfo.File_Date.Month.ToString("00") + "_" + videoInfo.File_Date.Day.ToString("00") + "/" + videoInfo.File_FullName;
    }

    private bool IsRecodeVideoExist(VideoInfo videoInfo)
    {
        bool isExist = false;
        if (videoInfo.File_FullName.IndexOf(".jnr") < 0)
        {
            return true;
        }
        WebRequest videoRecodeRequest = HttpWebRequest.Create(GetVideoFileUrl(videoInfo));
        WebResponse response = null;
        try
        {
            response = videoRecodeRequest.GetResponse();
            if (response.ContentLength > 0)
                isExist = true;
        }
        catch
        {
            isExist = false;
        }
        finally
        {
            if (videoRecodeRequest != null)
            {
                videoRecodeRequest = null;
            }
            if (response != null)
            {
                response.Close();
                response = null;
            }
        }
        return isExist;
    }

}