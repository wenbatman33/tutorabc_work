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
/// RatingInfo 的摘要描述
/// </summary>
/// 
[Serializable]
public class RatingInfo
{
    private ArrayList usersRatingList = new ArrayList();
    private bool isView = false;

    public bool IsView
    {
        get { return isView; }
        set { isView = value; }
    }
    public ArrayList UsersRatingList
    {
        get { return usersRatingList; }
        set { usersRatingList = value; }
    }
    private ArrayList usersRatingDate = new ArrayList();

    public ArrayList UsersRatingDate
    {
        get { return usersRatingDate; }
        set { usersRatingDate = value; }
    }
    private double userRating = double.MinValue;

    public double UserRating
    {
        get { return userRating; }
        set { userRating = value; }
    }
    private int viewCount = 0;
    private double avgRating = 0;

    public double AvgRating
    {
        get {
            double totalRating = 0;
            foreach (double aRating in UsersRatingList)
            {
                totalRating += aRating;
            }
            return totalRating / (double)UsersRatingList.Count;
        }

        set { avgRating = value; }
    }
    public int ViewCount
    {
        get { return viewCount; }
        set { viewCount = value; }
    }
        public RatingInfo()
        {

            //
            // TODO: 在此加入建構函式的程式碼
            //
        }

}
