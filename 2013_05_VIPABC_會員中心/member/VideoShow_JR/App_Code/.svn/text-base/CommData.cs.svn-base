using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Collections;
using System.Drawing;

/// <summary>
/// CommData 的摘要描述
/// </summary>
public class CommData
{
    public static ArrayList allSpecialSessionVideoInfo = new ArrayList();
    public static ArrayList allRegularlSessionVideoInfo = new ArrayList();
	public static ArrayList allFamousSessionVideoInfo = new ArrayList();
    public static ArrayList allPackageSessionVideoInfo = new ArrayList();
    public static ArrayList specialSessionList = new ArrayList();
    public static DateTime lastUpdateTime = new DateTime();
    public static int videoPerPage = 6 ;
    //public static int videoCount = 500;//似乎會有執行序的問題，我把此靜態變數改到DBOperation
    public static int unlimitDownloadSec = 10080;
	public CommData()
	{
		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
    /*public static ArrayList GetallSpecialSessionVideoInfo()
    {
        DBOperation dboperation = new DBOperation();
        ArrayList tempList = new ArrayList();
        string[] unValidRecodes = dboperation.GetUnValidSessionRecodes();
       
            foreach (VideoInfo aVideoInfo in allSpecialSessionVideoInfo)
            {
                bool isUnvalid = false;
                foreach (string aUnValidSessionRecode in unValidRecodes)
                {
                    if (aVideoInfo.SessionSN.Equals(aUnValidSessionRecode))
                        isUnvalid = true;
                }
                if(!isUnvalid)
                    tempList.Add(aVideoInfo.CloneVideoInfo());
            }
        
        return tempList;
    }
    public static ArrayList GetallRegularlSessionVideoInfo()
    {
        DBOperation dboperation = new DBOperation();
        ArrayList tempList = new ArrayList();
        string[] unValidRecodes = dboperation.GetUnValidSessionRecodes();
       
            foreach (VideoInfo aVideoInfo in allRegularlSessionVideoInfo)
            {
                    bool isUnvalid = false;
                    foreach (string aUnValidSessionRecode in unValidRecodes)
                    {
                        if (aVideoInfo.SessionSN.Equals(aUnValidSessionRecode))
                            isUnvalid = true;
                    }
                    if (!isUnvalid)
                        tempList.Add(aVideoInfo.CloneVideoInfo());
            }
        
        return tempList;
    }*/
    /// <summary>
    /// //PO add 
    /// </summary>
    /// <returns></returns>
    /*public static ArrayList GetallPackageSessionVideoInfo()
    {
        DBOperation dboperation = new DBOperation();
        ArrayList tempList = new ArrayList();
        string[] unValidRecodes = dboperation.GetUnValidSessionRecodes();
       
            foreach (VideoInfo aVideoInfo in allPackageSessionVideoInfo)
            {*/ /*
                bool isUnvalid = false;
                foreach (string aUnValidSessionRecode in unValidRecodes)
                {
                    if (aVideoInfo.SessionSN.Equals(aUnValidSessionRecode))
                        isUnvalid = true;
                }
                if (!isUnvalid)*/
                    /*tempList.Add(aVideoInfo.CloneVideoInfo());
                   
            }
            tempList.Sort(new SortVideoCompare());
        
        return tempList;
    }*/

	/*public static ArrayList GetallFamousSessionVideoInfo()
	{
		DBOperation dboperation = new DBOperation();
		ArrayList tempList = new ArrayList();
		string[] unValidRecodes = dboperation.GetUnValidSessionRecodes();
	
			foreach (VideoInfo aVideoInfo in allFamousSessionVideoInfo)
			{
				bool isUnvalid = false;
				foreach (string aUnValidSessionRecode in unValidRecodes)
				{
					if (aVideoInfo.SessionSN.Equals(aUnValidSessionRecode))
						isUnvalid = true;
				}
				if (!isUnvalid)
					tempList.Add(aVideoInfo.CloneVideoInfo());
			}
		
		return tempList;
	}*/
    public static string CreatePaddingString(string orgText, float maxLen)
    {

        float totalLen = 10.8671875f;
        char[] chars = orgText.ToCharArray();
        int tesxCount = 0;
        for (int stringIndex = 0; stringIndex < chars.Length; stringIndex++)
        {
            totalLen += GetTxetSize(chars[stringIndex].ToString());
            if (totalLen > maxLen)
                break;
            tesxCount++;
        }
        if (tesxCount == orgText.Length)
            return orgText;
        return orgText.Substring(0, tesxCount - 1) + "..";
    }
    public static float GetTxetSize(string aChar)
    {
        Bitmap bmp = new Bitmap(139, 37);
        Graphics g = Graphics.FromImage(bmp);
        SizeF size = g.MeasureString(aChar, new Font("Arial", 9, FontStyle.Bold));
        bmp.Dispose();
        g.Dispose();
        return size.Width;
    }
    public static SearchInfo setSearchInfo(string page, string sessionType, string interestType, string materialLevel, string searchString, string orderType)
    { 
        SearchInfo searchInfo = new SearchInfo();
        if (string.IsNullOrEmpty(page))
            page = "1";
        if (string.IsNullOrEmpty(sessionType))
            sessionType = "";
        if (string.IsNullOrEmpty(interestType))
            interestType = "";
        if (string.IsNullOrEmpty(searchString))
            searchString = "";
        if (string.IsNullOrEmpty(orderType))
            orderType = "rating";

        searchInfo.Page = page;
        searchInfo.VideoSessionType = sessionType;
        searchInfo.VideoInterestType = interestType;
        if (string.IsNullOrEmpty(materialLevel)) {
            materialLevel = "0";
        }
        searchInfo.VideoLevel = Convert.ToInt32(materialLevel);
        searchInfo.SearchPatthen = searchString;
        searchInfo.RankType = orderType;

        return searchInfo;
    }
}
public class SortVideoCompare : IComparer
{
    public int Compare(object x, object y)
    {
        VideoInfo objVideoA = (VideoInfo)x;
        VideoInfo objVideoB = (VideoInfo)y;
        return ((new CaseInsensitiveComparer()).Compare(objVideoA.Description, objVideoB.Description));
    }
}
