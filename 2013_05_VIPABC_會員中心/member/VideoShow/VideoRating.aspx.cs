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

public partial class VideoRating : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        if (Session["client_sn"] == null)
            Response.Redirect("http://www.vipabc.com");
        if (!IsPostBack)
        {
            Base64 base64 = new Base64();
            DBOperation dBOperation = new DBOperation();
            if (Request["sn"] == null)
            {
                Response.Redirect("http://www.vipabc.com");
                return;
            }
            if (Request["csn"] == null)
            {
                Response.Redirect("http://www.vipabc.com");
                return;
            }
            string sn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["sn"].ToString()), "ctabcpc07");
            string csn = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["csn"].ToString()), "ctabcpc07");
            string action = CryptLib.Crypt_Tool.DecryptAnyKeyStr(base64.DecodeFrom64(Request["action"].ToString()), "ctabcpc07");
            if (sn.Equals("") || csn.Equals("") || action.Equals(""))
            {
                Response.Redirect("http://www.vipabc.com");
                return;
            }
            dBOperation.ClientSn = csn;
            VideoInfo videoInfo = dBOperation.GetVideoInfoBySessionSN(sn);
            if (string.IsNullOrEmpty(videoInfo.SessionSN) || videoInfo.SessionSN.Equals(""))
            {
                Response.Redirect("http://www.vipabc.com");
                return;
            }
            if (action.Equals("viewAction"))
            {
                WebUserContro_RatingPanel1.SetVideoInfo(videoInfo, csn, true);
            }
            else
                WebUserContro_RatingPanel1.SetVideoInfo(videoInfo, csn, false);
        }
    }
}
