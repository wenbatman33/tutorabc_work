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
/// VideoInfo 的摘要描述
/// </summary>
[Serializable]
public class VideoInfo
{
    string course = "";

    public string Course
    {
        get { return course; }
        set { course = value; }
    }
    string nameOfVideo = "";
    private int sessionLevel = 0;

    public int SessionLevel
    {
        get { return sessionLevel; }
        set { sessionLevel = value; }
    }
    public string NameOfVideo
    {
        get { return nameOfVideo.Trim(); }
        set { nameOfVideo = value; }
    }
    string consultant = "";

    public string Consultant
    {
        get { return consultant.Trim(); }
        set { consultant = value; }
    }
    ArrayList lVList = new ArrayList();
    ArrayList topicList = new ArrayList();

    public ArrayList TopicList
    {
        get { return topicList; }
        set { topicList = value; }
    }
    public ArrayList LVList
    {
        get { return lVList; }
        set { lVList = value; }
    }

    ArrayList sessionStudent = new ArrayList();

    public ArrayList SessionStudent
    {
        get { return sessionStudent; }
        set { sessionStudent = value; }
    }
    private string thumbImage = "";

    public string ThumbImage
    {
        get { return thumbImage; }
        set { thumbImage = value; }
    }
    private string contentType = "";

    public string SessionType
    {
        get { return contentType; }
        set { contentType = value; }
    }
    private bool isOwnSession = false;

    public bool IsOwnSession
    {
        get { return isOwnSession; }
        set { isOwnSession = value; }
    }
    private string sessionSN = "";

    public string SessionSN
    {
        get { return sessionSN; }
        set { sessionSN = value; }
    }
    private double course_rating = Double.MinValue;

    public double SessionRating
    {
        get { return course_rating; }
        set { course_rating = value; }
    }
    private double mtl_Rating = 0;

    public double Mtl_Rating
    {
        get
        {
            double mt_Rating = 0;
            if (mtl_RatingIngo.UsersRatingList.Count > 0)
                return (SessionRating + mtl_RatingIngo.AvgRating) / 2.0;
            else
                return SessionRating; 
        }
        set { mtl_Rating = value; }
    }
    private string file_FullName = "";
    private RatingInfo mtl_RatingIngo = new RatingInfo();

    public RatingInfo Mtl_RatingInfo
    {
        get { return mtl_RatingIngo; }
        set { mtl_RatingIngo = value; }
    }


    public string File_FullName
    {
        get { return file_FullName; }
        set { file_FullName = value; }
    }
    private string file_SN = "";

    public string File_SN
    {
        get { return file_SN; }
        set { file_SN = value; }
    }
    private string file_Path = "";

    public string File_Path
    {
        get { return file_Path; }
        set { file_Path = value; }
    }
    private DateTime file_Date = new DateTime();

    public DateTime File_Date
    {
        get { return file_Date; }
        set { file_Date = value; }
    }
    private int focus_Grammar;

    public int Focus_Grammar
    {
        get { return focus_Grammar; }
        set { focus_Grammar = value; }
    }
    private int focus_Vocabulary;

    public int Focus_Vocabulary
    {
        get { return focus_Vocabulary; }
        set { focus_Vocabulary = value; }
    }
    private int focus_Pronuciation;

    public int Focus_Pronuciation
    {
        get { return focus_Pronuciation; }
        set { focus_Pronuciation = value; }
    }
    private int focus_Speaking;

    public int Focus_Speaking
    {
        get { return focus_Speaking; }
        set { focus_Speaking = value; }
    }
    private int focus_Listening;

    public int Focus_Listening
    {
        get { return focus_Listening; }
        set { focus_Listening = value; }
    }
    private int focus_Reading;

    public int Focus_Reading
    {
        get { return focus_Reading; }
        set { focus_Reading = value; }
    }
    private bool isClientUseThisSessionSn;

    public bool IsClientUseThisSessionSn
    {
        get { return isClientUseThisSessionSn; }
        set { isClientUseThisSessionSn = value; }
    }
    private bool isClientUseThisMaterial;

    public bool IsClientUseThisMaterial
    {
        get { return isClientUseThisMaterial; }
        set { isClientUseThisMaterial = value; }
    }
    private bool isClientUseThisRecordFile;

    public bool IsClientUseThisRecordFile
    {
        get { return isClientUseThisRecordFile; }
        set { isClientUseThisRecordFile = value; }
    }
  
    private string description = "";

    public string Description
    {
        get { return description; }
        set { description = value; }
    }
    private string consultantName = "";

    public string ConsultantName
    {
        get { return consultantName; }
        set { consultantName = value; }
    }
    public ArrayList Focus_List
    {
        get
        {
            int topFocusScore = GetTopFocusScire();
            ArrayList topfocusList = new ArrayList();
            if (this.Focus_Grammar == topFocusScore)
            {
                topfocusList.Add("Grammar");
            }
            if (this.Focus_Listening == topFocusScore)
            {
                topfocusList.Add("Listening");
            }
            if (this.Focus_Pronuciation == topFocusScore)
            {
                topfocusList.Add("Pronuciation");
            }
            if (this.Focus_Reading == topFocusScore)
            {
                topfocusList.Add("Reading");
            }
            if (this.Focus_Speaking == topFocusScore)
            {
                topfocusList.Add("Speaking");

            }
            if (this.Focus_Vocabulary == topFocusScore)
            {
                topfocusList.Add("Vocabulary");
            }
            return topfocusList;
        }
       
    }
    private int GetTopFocusScire()
    {
        ArrayList focusList = new ArrayList();
        focusList.Add(this.Focus_Grammar);
        focusList.Add(this.Focus_Listening);
        focusList.Add(this.Focus_Pronuciation);
        focusList.Add(this.Focus_Reading);
        focusList.Add(this.Focus_Speaking);
        focusList.Add(this.Focus_Vocabulary);
        focusList.Sort();
        int topFocusScore = Convert.ToInt32(focusList[focusList.Count - 1]);
        focusList.Clear();
        focusList = null;
        return topFocusScore;
    }
    string courseCount = "";

    public string CourseCount
    {
        get { return courseCount; }
        set { courseCount = value; }
    }
    public VideoInfo CloneVideoInfo()
    {
        VideoInfo videoInfo = new VideoInfo();
        videoInfo.File_SN = file_SN;
        videoInfo.Consultant = this.consultant;
        videoInfo.SessionType = this.contentType;
        videoInfo.Course = this.course;
        videoInfo.File_FullName = this.file_FullName;
        videoInfo.LVList = (ArrayList)this.lVList.Clone();
        videoInfo.NameOfVideo = this.nameOfVideo;
        videoInfo.SessionRating = this.course_rating;
        videoInfo.SessionSN = this.sessionSN;
        videoInfo.SessionStudent = (ArrayList)this.sessionStudent.Clone();
        videoInfo.SessionType = this.SessionType;
        videoInfo.ThumbImage = this.ThumbImage;
        videoInfo.TopicList = (ArrayList)this.topicList.Clone();
        videoInfo.Mtl_RatingInfo = this.Mtl_RatingInfo;
        videoInfo.File_Date = this.File_Date;
        videoInfo.File_Path = this.file_Path;
        videoInfo.Focus_Grammar = this.Focus_Grammar;
        videoInfo.Focus_Listening = this.Focus_Listening;
        videoInfo.Focus_Pronuciation = this.Focus_Pronuciation;
        videoInfo.Focus_Reading = this.Focus_Reading;
        videoInfo.Focus_Speaking = this.Focus_Speaking;
        videoInfo.Focus_Vocabulary = this.Focus_Vocabulary;
        videoInfo.description = this.description;
        videoInfo.ConsultantName = this.consultantName;
        videoInfo.courseCount = this.courseCount;
        return videoInfo;
    }
}
