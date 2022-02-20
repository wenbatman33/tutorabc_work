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

public partial class UI_WebUserControl_ShowRateing : System.Web.UI.UserControl
{
    private enum StarType
    {
        Full,
        Half,
        Blank,
    }

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    public void SetRating(float rating)
    {
        int full_ShowCount = (int)(rating * 10)/20;
        int half_ShowCount =-1;
        int half_Point = ((int)(rating * 10)-(full_ShowCount*20) ) % 20;
        if (half_Point >= 10)
            half_ShowCount = full_ShowCount+1;
        for (int fullIndex = 1; fullIndex <= full_ShowCount; fullIndex++)
        {
            ShowRatingImge(fullIndex, StarType.Full);
        }
        if (half_ShowCount != -1)
            ShowRatingImge(half_ShowCount, StarType.Half);
        else
            half_ShowCount = full_ShowCount;
        for (int BlankIndex = (half_ShowCount + 1); BlankIndex <= 5; BlankIndex++)
        {
            ShowRatingImge(BlankIndex, StarType.Blank);
        }
        Panel1.ToolTip = "評比:"+rating.ToString();
 
    }

    private void ShowRatingImge(int starCount, StarType starType)
    {
        foreach (Control aControl in Panel1.Controls)
        {
            if (aControl is Image)
            {
                Image tempImage = (Image)aControl;
                if (tempImage.ID.Equals("Image_Star" + starCount))
                {
                    tempImage.ImageUrl = "../Image/star_" + starType + ".gif";
                }
            }
        }
    }

}
