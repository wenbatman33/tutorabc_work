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

/// <summary>
/// SearchInfo 的摘要描述
/// </summary>
[Serializable]
public class SearchInfo
{
    private string videoInterestType = "";

    public string VideoInterestType
    {
        get { return videoInterestType; }
        set { videoInterestType = value; }
    }
    private int videoLevel = 0;

    public int VideoLevel
    {
        get { return videoLevel; }
        set { videoLevel = value; }
    }
    private string consultaneName = "";

    public string ConsultaneName
    {
        get { return consultaneName; }
        set { consultaneName = value; }
    }
    private string searchPatten = "";

    public string SearchPatthen
    {
        get { return searchPatten; }
        set { searchPatten = value; }
    }
    private string videoSessionType = "";

    public string VideoSessionType
    {
        get { return videoSessionType; }
        set { videoSessionType = value; }
    }
    private string rankType = "rating";

    public string RankType
    {
        get { return rankType; }
        set { rankType = value; }
    }
    private string fous_Type = "";
    public string Fous_Type
    {
        get { return fous_Type; }
        set { fous_Type = value; }
    }
    private string page = "1";
    public string Page
    {
        get { return page; }
        set { page = value; }
    }
	public SearchInfo()
	{
        

		//
		// TODO: 在此加入建構函式的程式碼
		//
	}
}
