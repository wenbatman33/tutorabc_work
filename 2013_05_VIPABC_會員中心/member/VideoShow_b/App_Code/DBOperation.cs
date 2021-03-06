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
using System.Data.SqlClient;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;


/// <summary>
/// DBOperation 的摘要描述
/// </summary>
public class DBOperation
{
    private SearchInfo clientSearchInfo = new SearchInfo();
    private int intSelectLength = 0;

    public SearchInfo ClientSearchInfo
    {
        get { return clientSearchInfo; }
        set { clientSearchInfo = value; }
    }
    //基本上，只要設置過一次clientSn就可以將此變數一直存放，畢竟一次只會有一個客戶查詢
   private static string  clientSn = "0";
    
    public string ClientSn
    {
        get { return clientSn; }
        set { clientSn = value; }
    }

	public DBOperation()
	{
        
	}

    /*public void UpdateVideoSelectedInfo()
    {
        VideoInfo[] ssVideo = SpecialSessioVideoList();
        VideoInfo[] rgVideo = AllRegularVideoList();
        VideoInfo[] gsVideo = PacketSessionVideoList();//Po_2010 add
        WriteToSelectedTable(ssVideo);
        WriteToSelectedTable(rgVideo);
        WriteToSelectedTable(gsVideo);//Po_2010 add
    }*/

    private void WriteToSelectedTable(VideoInfo[] videos)
    {
        foreach (VideoInfo aVideoInfo in videos)
        {
            if (IsExistSelectedTable(aVideoInfo))
            {
                UpdateSelectedTable(aVideoInfo);
            }
            else
            {
                InsertSelectedTable(aVideoInfo);
            }
        }
    }

    private string ArrayListtoString(ArrayList list)
    {
        StringBuilder sb = new StringBuilder();
        try
        {
            foreach (string aData in list)
            {
                sb.Append(aData + ",");
            }
            if (sb.ToString().Length > 0)
            {
                return sb.ToString().Substring(0, sb.ToString().Length - 1);
            }
            return sb.ToString();
        }
        finally
        {
            sb = null;
        }
    } 

    private bool InsertSelectedTable(VideoInfo videoInfo)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " INSERT INTO SessionRecord_selected_by_material(course,session_sn,client_sn_list,File_Full_name,ltitle,ltopic,rating,consultant,attend_level,session_type,file_sn,file_path,file_date) " +
                               " VALUES(@course,@session_sn,@client_sn_list,@File_Full_name,@ltitle,@ltopic,@rating,@consultant,@attend_level,@session_type,@file_sn,@file_path,@file_date) ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@course", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@client_sn_list", SqlDbType.VarChar);
            cmd.Parameters.Add("@File_Full_name", SqlDbType.Text);
            cmd.Parameters.Add("@ltitle", SqlDbType.VarChar);
            cmd.Parameters.Add("@ltopic", SqlDbType.VarChar);
            cmd.Parameters.Add("@rating", SqlDbType.VarChar);
            cmd.Parameters.Add("@consultant", SqlDbType.VarChar);
            cmd.Parameters.Add("@attend_level", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_type", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_path", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_date", SqlDbType.DateTime);
            cmd.Parameters["@course"].Value = videoInfo.Course;
            cmd.Parameters["@session_sn"].Value = videoInfo.SessionSN;
            cmd.Parameters["@client_sn_list"].Value = ArrayListtoString(videoInfo.SessionStudent);
            cmd.Parameters["@File_Full_name"].Value = videoInfo.File_FullName;
            cmd.Parameters["@ltitle"].Value = videoInfo.NameOfVideo;
            cmd.Parameters["@ltopic"].Value = ArrayListtoString(videoInfo.TopicList);
            cmd.Parameters["@rating"].Value = videoInfo.SessionRating.ToString();
            cmd.Parameters["@consultant"].Value = videoInfo.Consultant;
            cmd.Parameters["@attend_level"].Value = ArrayListtoString(videoInfo.LVList);
            cmd.Parameters["@session_type"].Value = videoInfo.SessionType;
            cmd.Parameters["@file_sn"].Value = videoInfo.File_SN;
            cmd.Parameters["@file_path"].Value = videoInfo.File_Path;
            cmd.Parameters["@file_date"].Value = videoInfo.File_Date;
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public bool IsRecodeValid(VideoInfo videoInfo)
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT Session_Record_Number " +
                               " FROM SessionRecord_fileinfo " +
                               " WHERE (Session_Record_Number = @Session_Record_Number) AND (File_show_view_state = 'Y') AND (File_show_view_state = 'Y') ";//...同一個條件是怎了=_=
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Session_Record_Number", SqlDbType.VarChar);
            cmd.Parameters["@Session_Record_Number"].Value = videoInfo.SessionSN;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    public string[] GetUnValidSessionRecodes()
    {
        ArrayList unValidRecodeList = new ArrayList();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SessionRecord_selected_by_material.ltitle, SessionRecord_fileinfo.File_view_state, SessionRecord_fileinfo.File_show_view_state, "+ 
                               " SessionRecord_selected_by_material.session_sn, SessionRecord_selected_by_material.course "+
                               " FROM SessionRecord_selected_by_material INNER JOIN "+
                               " SessionRecord_fileinfo ON SessionRecord_selected_by_material.session_sn = SessionRecord_fileinfo.Session_Record_Number "+
                               " WHERE (SessionRecord_fileinfo.File_view_state = 'N') AND (SessionRecord_fileinfo.File_show_view_state = 'N') ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            while (rs.Read())
            {
                unValidRecodeList.Add(rs["session_sn"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return (string[])unValidRecodeList.ToArray(typeof(string));
    }

    private bool UpdateSelectedTable(VideoInfo videoInfo)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " UPDATE SessionRecord_selected_by_material SET session_sn = @session_sn,client_sn_list = @client_sn_list,File_Full_name = @File_Full_name,ltitle = @ltitle,ltopic = @ltopic,rating=@rating,consultant = @consultant,attend_level = @attend_level,file_sn = @file_sn,file_path = @file_path,file_date = @file_date" +
                               " WHERE course = @course AND session_type =@session_type";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@course", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@client_sn_list", SqlDbType.VarChar);
            cmd.Parameters.Add("@File_Full_name", SqlDbType.VarChar);
            cmd.Parameters.Add("@ltitle", SqlDbType.VarChar);
            cmd.Parameters.Add("@ltopic", SqlDbType.VarChar);
            cmd.Parameters.Add("@rating", SqlDbType.VarChar);
            cmd.Parameters.Add("@consultant", SqlDbType.VarChar);
            cmd.Parameters.Add("@attend_level", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_type", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_path", SqlDbType.VarChar);
            cmd.Parameters.Add("@file_date", SqlDbType.DateTime);
            cmd.Parameters["@course"].Value = videoInfo.Course;
            cmd.Parameters["@session_sn"].Value = videoInfo.SessionSN;
            cmd.Parameters["@client_sn_list"].Value = ArrayListtoString(videoInfo.SessionStudent);
            cmd.Parameters["@File_Full_name"].Value = videoInfo.File_FullName;
            cmd.Parameters["@ltitle"].Value = videoInfo.NameOfVideo;
            cmd.Parameters["@ltopic"].Value = ArrayListtoString(videoInfo.TopicList);
            cmd.Parameters["@rating"].Value = videoInfo.SessionRating;
            cmd.Parameters["@consultant"].Value = videoInfo.Consultant;
            cmd.Parameters["@attend_level"].Value = ArrayListtoString(videoInfo.LVList);
            cmd.Parameters["@session_type"].Value = videoInfo.SessionType;
            cmd.Parameters["@file_sn"].Value = videoInfo.File_SN;
            cmd.Parameters["@file_path"].Value = videoInfo.File_Path;
            cmd.Parameters["@file_date"].Value = videoInfo.File_Date;
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    private bool IsExistSelectedTable(VideoInfo videoInfo)
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT course FROM SessionRecord_selected_by_material where course='" + videoInfo.Course + "' AND session_type='" + videoInfo.SessionType + "' ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);

            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                return true;
            }
        }
        catch (Exception ex)
        {

        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    //好像沒在用
    private VideoInfo[] SpecialSessioVideoList()
    {
        ArrayList resultList = new ArrayList();
        ArrayList mtlList = new ArrayList();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = @"
                            SELECT * FROM dbo.vSessionVideosVIPABC WITH(NOLOCK) WHERE session_type='GS'
                        ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return null;
            while (rs.Read())
            {
                //if (mtlList.IndexOf(rs["course"].ToString()) == -1)
                resultList.Add(rs["course"].ToString());               
            }
            /*foreach (string mtl in mtlList)
            {
                VideoInfo videoInfo = new VideoInfo();
                videoInfo = GetTop1SessionVideoByMtl(mtl, true);
                if (videoInfo == null)
                    continue;
                videoInfo.SessionType = "GS";
                videoInformation.Add(GetVideoInfoFromDataReader(rs));
                resultList.Add(videoInfo);

            }*/
            return (VideoInfo[])resultList.ToArray(typeof(VideoInfo));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
            resultList.Clear();
            resultList = null;
            mtlList.Clear();
            mtlList = null;
        }
        return null;  
        
    }

    /*public VideoInfo[] PacketSessionVideoList()//Po_2010 06/14 add
    {

        ReadFromSelectedTable();
        ArrayList resultList = new ArrayList();
        try
        {
            foreach (VideoInfo aVideoInfo in CommData.GetallPackageSessionVideoInfo())
            {
                resultList.Add(aVideoInfo.CloneVideoInfo());
            }
            return FilterBySearchInfo(resultList, true);
        }
        finally
        {
            resultList.Clear();
            resultList = null;
        }

    }*/

    /*private VideoInfo[] AllRegularVideoList()
    {
        ArrayList resultList = new ArrayList();
        Hashtable videoTable = new Hashtable();
        ArrayList mtlList = new ArrayList();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT valid, course " +
                               " FROM material " +
                               " WHERE (valid = 1) ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return null;
            while (rs.Read())
            {
                mtlList.Add(rs["course"].ToString());
            }
            foreach (string mtl in mtlList)
            {
                VideoInfo videoInfo = new VideoInfo();
                videoInfo = GetTop1SessionVideoByMtl(mtl, false);
                if (videoInfo == null)
                    continue;
                videoInfo.SessionType = "RG";
                resultList.Add(videoInfo);
            }
            return (VideoInfo[])resultList.ToArray(typeof(VideoInfo));

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
            resultList.Clear();
            resultList = null;
            mtlList.Clear();
            mtlList = null;
        }
        return null;
    }*/

    /*private VideoInfo GetTop1SessionVideoByMtl(string mtlNumber,bool isSpecialSession)
    {
        Hashtable videoTable = new Hashtable();
  
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT derivedtbl_1.attend_mtl_1 AS mtl, SUM(derivedtbl_1.fbv) AS fbv, COUNT(derivedtbl_1.sn) AS [user], derivedtbl_1.session_sn,"+
                                "(SELECT CAST(LEFT(ROUND((SUM(client_session_evaluation.consultant_points) + 0.0) / COUNT(client_session_evaluation.sn), 2), 4) AS float)"+
                                "AS rating"+
                                    " FROM client_session_evaluation LEFT OUTER JOIN "+
                                            "client_attend_list ON client_session_evaluation.client_sn = client_attend_list.client_sn AND "+
                                            "client_session_evaluation.session_sn = client_attend_list.session_sn"+
                                    " WHERE (client_attend_list.attend_consultant = derivedtbl_1.attend_consultant)"+
                                    " GROUP BY client_attend_list.attend_consultant) AS rating, SessionRecord_fileinfo.File_view_state,"+
                                            "con_basic.basic_fname + ' ' + con_basic.basic_lname AS consultant, derivedtbl_1.attend_consultant, derivedtbl_1.client_sn, material.ltitle,"+
                                            "derivedtbl_1.attend_level, material.ltopic, SessionRecord_fileinfo.File_Full_name,SessionRecord_fileinfo.SN,SessionRecord_fileinfo.file_path,SessionRecord_fileinfo.file_datatime" +
                                " FROM (SELECT client_attend_list_1.sn, client_attend_list_1.client_sn, client_attend_list_1.attend_date, client_attend_list_1.attend_sestime,"+
                                                "client_attend_list_1.attend_livesession_types, client_attend_list_1.attend_consultant, client_attend_list_1.attend_datetime,"+
                                                "client_attend_list_1.attend_room, client_attend_list_1.attend_mtl_1, client_attend_list_1.attend_mtl_2, client_attend_list_1.attend_remind,"+
                                                "client_attend_list_1.attend_approve, client_attend_list_1.attend_level, client_attend_list_1.chanel_iP, client_attend_list_1.chanel_datetime,"+
                                                "client_attend_list_1.session_sn, client_attend_list_1.attend_type, client_attend_list_1.valid, client_attend_list_1.state,"+
                                                "client_attend_list_1.class_yesno, client_attend_list_1.block_classmate, client_attend_list_1.block_consultant, client_attend_list_1.rand_str,"+
                                                "client_attend_list_1.session_state, client_attend_list_1.session_remark, client_attend_list_1.refund, client_attend_list_1.view_ip,"+
                                                "client_attend_list_1.session_room_view_ip, client_attend_list_1.session_room_state, client_attend_list_1.client_online_state,"+
                                                "client_attend_list_1.client_online_state_date, client_attend_list_1.original_con, client_attend_list_1.org_date,"+
                                                "client_attend_list_1.joinnet_ver, client_attend_list_1.session_country, client_attend_list_1.join_time, client_attend_list_1.unsummary,"+
                                                "client_attend_list_1.special_sn, ISNULL(client_session_evaluation_1.consultant_points, 0)"+
                                                "+ ISNULL(client_session_evaluation_1.materials_points, 0) + ISNULL(client_session_evaluation_1.overall_points, 0) AS fbv"+
                                " FROM client_attend_list AS client_attend_list_1 LEFT OUTER JOIN "+
                                                "client_session_evaluation AS client_session_evaluation_1 ON "+
                                                "client_attend_list_1.session_sn = client_session_evaluation_1.session_sn AND "+
                                                "client_attend_list_1.client_sn = client_session_evaluation_1.client_sn"+
                                " WHERE (client_attend_list_1.attend_mtl_1 IS NOT NULL) AND (client_attend_list_1.session_sn IS NOT NULL) AND (client_attend_list_1.valid = 1))"+
                                                " AS derivedtbl_1 INNER JOIN "+
                                                "SessionRecord_fileinfo ON derivedtbl_1.session_sn = SessionRecord_fileinfo.Session_Record_Number INNER JOIN "+
                                                "con_basic ON derivedtbl_1.attend_consultant = con_basic.con_sn INNER JOIN "+
                                                "material ON derivedtbl_1.attend_mtl_1 = material.course"+
                                " WHERE (derivedtbl_1.attend_mtl_1 = '" + mtlNumber + "')" +
                                " GROUP BY derivedtbl_1.session_sn, derivedtbl_1.attend_mtl_1, SessionRecord_fileinfo.File_view_state, con_basic.basic_fname + ' ' + con_basic.basic_lname,"+
                                "derivedtbl_1.attend_consultant, derivedtbl_1.client_sn, material.ltitle, derivedtbl_1.attend_level, material.ltopic,SessionRecord_fileinfo.File_Full_name,SessionRecord_fileinfo.SN,SessionRecord_fileinfo.file_path,SessionRecord_fileinfo.file_datatime" +
                                " HAVING (SessionRecord_fileinfo.File_view_state = 'Y')"+
                                " ORDER BY fbv DESC, [user], rating DESC";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return null;
            while (rs.Read())
            {
                if (rs["SN"] == DBNull.Value || rs["file_datatime"] == DBNull.Value || rs["rating"] == DBNull.Value )
                    continue;
               if(videoTable[rs["session_sn"].ToString()] != null)
               {
                   VideoInfo tempVideoInfo = (VideoInfo)(videoTable[rs["session_sn"].ToString()]);
                  　tempVideoInfo.SessionStudent.Add(rs["client_sn"].ToString());
                   if (tempVideoInfo.LVList.IndexOf(rs["attend_level"].ToString()) == -1)
                   {
                       tempVideoInfo.LVList.Add(rs["attend_level"].ToString());
                   }
               }
               else
               {
                   VideoInfo videoInfo = new VideoInfo();
                   videoInfo.SessionStudent.Clear();
                   videoInfo.SessionRating = Convert.ToDouble(rs["rating"].ToString());
                   videoInfo.NameOfVideo = rs["ltitle"].ToString();
                   videoInfo.SessionSN = rs["session_sn"].ToString();
                   videoInfo.TopicList = GetTopicList(rs["ltopic"].ToString());
                   videoInfo.SessionStudent.Add(rs["client_sn"].ToString());
                   videoInfo.LVList.Add(rs["attend_level"].ToString());
                   videoInfo.Consultant = rs["consultant"].ToString();
                   videoInfo.Course = rs["mtl"].ToString();
                   videoInfo.File_FullName = rs["File_Full_name"].ToString();
                   videoInfo.File_SN = rs["SN"].ToString();
                   videoInfo.File_Date = DateTime.Parse(rs["file_datatime"].ToString());
                   videoInfo.File_Path = rs["file_path"].ToString();
                   videoTable.Add(rs["session_sn"].ToString(), videoInfo);

               }
            
            }
            return GetTopMaterialVideo(videoTable, isSpecialSession);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
            videoTable.Clear();
            videoTable = null;
        }
        return null;
    }*/

    /*private VideoInfo GetTopMaterialVideo(Hashtable videoTable ,bool isSpecialSession)
    {
        Hashtable tempVideoTable = new Hashtable();
        try
        {
            tempVideoTable = SpecialSessionFilter(videoTable, isSpecialSession);
            if (isSpecialSession)
            {
                tempVideoTable = SessionRatingFilter(tempVideoTable);
                tempVideoTable = StudentCountFilter(tempVideoTable);
                tempVideoTable = ConsultantRatingFilter(tempVideoTable);
            }
            else
            {
                tempVideoTable = LevelFilter(tempVideoTable);
                tempVideoTable = SessionRatingFilter(tempVideoTable);
                tempVideoTable = StudentCountFilter(tempVideoTable);
                tempVideoTable = ConsultantRatingFilter(tempVideoTable);
            }
            foreach (string aSession in tempVideoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)tempVideoTable[aSession];

                return tempVideoInfo;
            }
        }
        finally
        {
            tempVideoTable.Clear();
            tempVideoTable = null;
        }
        return null;
    }*/
  
    /*private Hashtable SpecialSessionFilter(Hashtable videoTable,bool isSpecialSession)
    {
        lock (CommData.specialSessionList)
        {
            Hashtable tempVideoTable = new Hashtable();
            if (CommData.specialSessionList.Count < 1)
            {
                SqlConnection conn = null;
                SqlCommand cmd = null;
                SqlDataReader rs = null;
                try
                {
                    string sql = " SELECT session_sn" +
                                       " FROM special_session_record";

                    conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
                    conn.Open();
                    cmd = new SqlCommand(sql, conn);
                    rs = cmd.ExecuteReader();
                    if (rs == null)
                        return null;
                    while (rs.Read())
                    {
                        CommData.specialSessionList.Add(rs["session_sn"].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (rs != null)
                        rs.Close();
                    if (cmd != null)
                        cmd.Dispose();
                    if (conn != null)
                        conn.Dispose();
                }
            }
            if (isSpecialSession)
            {
                foreach (string aSession in videoTable.Keys)
                {
                    if (CommData.specialSessionList.IndexOf(aSession) != -1)
                        tempVideoTable.Add(aSession, videoTable[aSession]);
                }
            }
            else
            {
                foreach (string aSession in videoTable.Keys)
                {
                    if (CommData.specialSessionList.IndexOf(aSession) == -1)
                        tempVideoTable.Add(aSession, videoTable[aSession]);
                }
            }
            return tempVideoTable;
        }
    }*/
    /*
    private Hashtable LevelFilter(Hashtable videoTable)
    {
        Hashtable tempVideoTable = new Hashtable();
        try
        {
            foreach (string aSession in videoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)videoTable[aSession];
                tempVideoInfo.LVList.Sort();
                if (tempVideoInfo.LVList[0].Equals(tempVideoInfo.LVList[tempVideoInfo.LVList.Count - 1]))
                {
                    int lev = Convert.ToInt32(tempVideoInfo.LVList[0]);
                    tempVideoInfo.LVList.Clear();
                    tempVideoInfo.LVList.Add(lev);
                    tempVideoTable.Add(aSession, tempVideoInfo);
                }
            }
            return tempVideoTable;
        }
        finally
        {
            tempVideoTable.Clear();
            tempVideoTable = null;
        }
    }

    private Hashtable SessionRatingFilter(Hashtable videoTable)
    {
        Hashtable tempVideoTable = new Hashtable();
        try
        {
            double maxRating = 0;
            foreach (string aSession in videoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)videoTable[aSession];
                if (tempVideoInfo.SessionRating > maxRating)
                {
                    maxRating = tempVideoInfo.SessionRating;
                }
            }
            foreach (string aSession in videoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)videoTable[aSession];
                if (tempVideoInfo.SessionRating == maxRating)
                {
                    tempVideoTable.Add(aSession, tempVideoInfo);
                }
            }
            return tempVideoTable;
        }
        finally
        {
            tempVideoTable.Clear();
            tempVideoTable = null;
        }
 
    }
    private Hashtable StudentCountFilter(Hashtable videoTable)
    {
        Hashtable tempVideoTable = new Hashtable();
        try
        {
            int minUserCount = Int32.MaxValue;
            foreach (string aSession in videoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)videoTable[aSession];
                if (tempVideoInfo.SessionStudent.Count < minUserCount)
                {
                    minUserCount = tempVideoInfo.SessionStudent.Count;
                }
            }
            foreach (string aSession in videoTable.Keys)
            {
                VideoInfo tempVideoInfo = (VideoInfo)videoTable[aSession];
                if (tempVideoInfo.SessionStudent.Count == minUserCount)
                {
                    tempVideoTable.Add(aSession, tempVideoInfo);
                }
            }
            return tempVideoTable;
        }
        finally
        {
            tempVideoTable.Clear();
            tempVideoTable = null;
        }
    }
    
    private Hashtable ConsultantRatingFilter(Hashtable videoTable)
    {
        return videoTable;
    }
    
    private ArrayList GetTopicList(string topicString)
    {
        ArrayList topicList = new ArrayList();
        try
        {
            topicString = topicString.Replace("--", "-");
            string[] topicDatas = topicString.Split('-');
            foreach (string aTopic in topicDatas)
            {
                if (!aTopic.Equals(""))
                    topicList.Add(aTopic);
            }
            return topicList;
        }
        finally
        {
            topicList.Clear();
            topicList = null;
        }
    }
    */
    public bool IsPointEnough(double needPoint)
    {
        double contactPoint = GetContractPoint();
        double usedPoint = GetUsePointPoint();
        if (((contactPoint - usedPoint) - needPoint) >= 0)
            return true;
        return false;
    }

    private double GetContractPoint()
    {
        double contactPoint = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT CAST((fp + rp) / 65 AS int) AS tp, sdate, edate FROM (SELECT SUM(add_points) AS fp, SUM(access_points) AS rp, MIN(product_sdate) AS sdate, MAX(product_edate) AS edate " +
                               " FROM dbo.client_purchase " +
                               " WHERE (client_sn ="+ clientSn+") AND (Valid = 1)) AS DERIVEDTBL ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return 0;
            while (rs.Read())
            {
                contactPoint += Convert.ToDouble(rs["tp"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return contactPoint;
    }

    private double GetUsePointPoint()
    {
        double usePoint = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SUM(pused) / 65 AS sused "+
                               " FROM (SELECT (CASE dbo.client_attend_list.refund WHEN 1 THEN 0 ELSE ((isnull(client_attend_list.svalue, 0) + 0.00) * 65) END) AS pused "+
                               " FROM dbo.client_attend_list LEFT OUTER JOIN "+
                               " dbo.cfg_live_session_points ON dbo.client_attend_list.attend_livesession_types = dbo.cfg_live_session_points.sn RIGHT OUTER JOIN "+
                               " (SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate "+
                               " FROM dbo.client_purchase "+
                               " WHERE (client_sn = @client_sn) AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) AND "+
                               " (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111)) AND (Valid = 1) "+
                               " GROUP BY client_sn) AS derivedtbl_1 ON CONVERT(varchar, dbo.client_attend_list.attend_date, 111) <= derivedtbl_1.edate AND "+
                               " CONVERT(varchar, dbo.client_attend_list.attend_date, 111) >= derivedtbl_1.sdate AND "+
                               " dbo.client_attend_list.client_sn = derivedtbl_1.client_sn "+
                               " WHERE (dbo.client_attend_list.valid = 1) AND (dbo.client_attend_list.refund IN (0, 1)) "+
                               " UNION ALL "+
                               " SELECT dbo.sessionrecord_view_list.view_cost AS pused "+
                               " FROM dbo.sessionrecord_view_list RIGHT OUTER JOIN "+
                               " (SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate "+
                               " FROM dbo.client_purchase AS client_purchase_1 "+
                               " WHERE (client_sn = @client_sn) AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) AND " +
                               " (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111)) AND (Valid = 1) "+
                               " GROUP BY client_sn) AS derivedtbl_1_1 ON CONVERT(varchar, dbo.sessionrecord_view_list.view_datetime, 111) "+
                               " <= derivedtbl_1_1.edate AND CONVERT(varchar, dbo.sessionrecord_view_list.view_datetime, 111) >= derivedtbl_1_1.sdate AND "+
                               " dbo.sessionrecord_view_list.client_sn = derivedtbl_1_1.client_sn "+
                               " WHERE (dbo.sessionrecord_view_list.client_sn IS NOT NULL)) AS a ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = clientSn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return 0;
            while (rs.Read())
            {
                if(rs["sused"]!=DBNull.Value)
                    usePoint += Convert.ToDouble(rs["sused"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return usePoint;
    }

    public bool IsFreeTimeDonlwoad()
    {
        DateTime contactDate = new DateTime();
        DateTime freeContactDate = new DateTime(2007,4,16);
        int currentHour = DateTime.Now.Hour;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT top 1 product_sn, product_sdate, product_edate FROM client_purchase WHERE (client_sn = @client_sn) AND (Valid = 1)  order by product_sn desc ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());

            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = clientSn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                contactDate = DateTime.Parse(rs["product_sdate"].ToString());
            }
        }
        catch (Exception ex)
        {
           return false;
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        if (contactDate <= freeContactDate && currentHour >=0 && currentHour < 10)
            return true;
        return false;
    }

    public bool IsContractValid()
    {
        DateTime contactDate = DateTime.MinValue;
        int currentHour = DateTime.Now.Hour;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT product_sn, product_sdate, product_edate FROM client_purchase WHERE (client_sn = @client_sn) AND (Valid = 1)  order by product_sn desc ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());

            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = clientSn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            while (rs.Read())
            {
                DateTime tempDate = DateTime.Parse(rs["product_edate"].ToString());
                if (tempDate > contactDate)
                {
                    contactDate = tempDate;
                }
            }
        }
        catch (Exception ex)
        {
            return false;
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        if (DateTime.Parse(DateTime.Now.ToShortDateString()) <= contactDate)
            return true;
        return false;
    }

    /*private void ReadFromSelectedTable()
    {
        if ( DateTime.Now.Subtract(CommData.lastUpdateTime).Days > 1 || CommData.GetallRegularlSessionVideoInfo().Count == 0 || CommData.GetallSpecialSessionVideoInfo().Count == 0 || CommData.GetallPackageSessionVideoInfo().Count == 0)
        {
            //他每次都會重讀，但又不清0.0...
            CommData.allFamousSessionVideoInfo.Clear();
            CommData.allPackageSessionVideoInfo.Clear();
            CommData.allRegularlSessionVideoInfo.Clear();
            CommData.allSpecialSessionVideoInfo.Clear();
            CommData.lastUpdateTime = DateTime.Now;
            SqlConnection conn = null;
            SqlCommand cmd = null;
            SqlDataReader rs = null;
            ArrayList rgList = new ArrayList();
            ArrayList ssList = new ArrayList();
            ArrayList fsList = new ArrayList();
            ArrayList gsList = new ArrayList();
            try
            {
                string sql = @" SELECT * FROM vSessionVideosVIPABC WITH(NOLOCK) ORDER BY rating";
                conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
                conn.Open();
                cmd = new SqlCommand(sql, conn);
                rs = cmd.ExecuteReader();
                if (rs == null)
                    return;

                while (rs.Read())
                {
                    VideoInfo videoInfo = GetVideoInfoFromDataReader(rs);
                    if (videoInfo != null)
                    {
                        if (rs["session_type"].ToString().Equals("RG"))
                        {
                            if (rgList.IndexOf(videoInfo) == -1)
                                rgList.Add(videoInfo.CloneVideoInfo());
                        }
                        if (rs["session_type"].ToString().Equals("GS"))//Po_2010 6.14 add
                        {
                            if (gsList.IndexOf(videoInfo) == -1)
                                gsList.Add(videoInfo.CloneVideoInfo());
                        }
                        if (rs["session_type"].ToString().Equals("SS"))
                        {

                            if (ssList.IndexOf(videoInfo) == -1)
                                ssList.Add(videoInfo.CloneVideoInfo());

                        }
                        if (rs["session_type"].ToString().Equals("FS"))
                        {

                            if (fsList.IndexOf(videoInfo) == -1)
                            {
                                videoInfo.IsOwnSession = true;
                                fsList.Add(videoInfo.CloneVideoInfo());
                            }

                        }
                    }
                }
                CommData.allRegularlSessionVideoInfo = rgList;
                CommData.allSpecialSessionVideoInfo = ssList;
                CommData.allPackageSessionVideoInfo = gsList;//Po_2010 add
                CommData.allFamousSessionVideoInfo = fsList;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (rs != null)
                    rs.Close();
                if (cmd != null)
                    cmd.Dispose();
                if (conn != null)
                    conn.Dispose();

                //CommData.allRegularlSessionVideoInfo.Sort(new myHighRatingClass());
                //CommData.allSpecialSessionVideoInfo.Sort(new myHighRatingClass());
                //CommData.allPackageSessionVideoInfo.Sort(new myHighRatingClass());
                rgList = null;
                ssList = null;
                gsList = null;
            }
        }
    }*/
    /*Lily*/
    public List<VideoInfo> SearchPatternSelectTable(SearchInfo searchInfo)
    {
        List<VideoInfo> videoInformation = new List<VideoInfo>();        
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        int count = 0;
        if (string.IsNullOrEmpty(searchInfo.Page))
            searchInfo.Page = "1";
        string sessionType = searchInfo.VideoSessionType;
        string interestType = searchInfo.VideoInterestType;
        string materialLevel = "";
        if(searchInfo.VideoLevel!=0)
             materialLevel = searchInfo.VideoLevel.ToString();
        string searchString = searchInfo.SearchPatthen;
        string orderType = searchInfo.RankType;
        int intPage = Convert.ToInt32(searchInfo.Page);
        try
        {
            string sql = "";
            if (string.IsNullOrEmpty(orderType)||orderType=="") {
                orderType = "rating";
            }
            sql = @"
                    WITH    materialPage
                              AS ( SELECT   dbo.vSessionVideosVIPABC.*
                                   FROM     vSessionVideosVIPABC WITH ( NOLOCK )
                                   INNER JOIN SessionRecord_fileinfo WITH ( NOLOCK ) ON vSessionVideosVIPABC.session_sn = SessionRecord_fileinfo.Session_Record_Number
                                   WHERE    course IS NOT NULL
                                   AND ((SessionRecord_fileinfo.File_view_state = 'Y') OR (SessionRecord_fileinfo.File_show_view_state = 'Y') )
            ";
            if(!string.IsNullOrEmpty(interestType)&&interestType!="")
            {
                sql += @"                   AND ltopic LIKE '%,'+@ltopic+',%'";
            }
            if(!string.IsNullOrEmpty(materialLevel)&&materialLevel!="")
            {
                sql += @"                   AND attend_level LIKE '%,'+@attend_level+',%'";
            }
            if(!string.IsNullOrEmpty(sessionType)&&sessionType!="")
            {
                sql += @"                   AND session_type = @sessionType";
            }
            if(!string.IsNullOrEmpty(searchString)&&searchString!="")
            {
                sql += @"                   AND ( Replace(consultant, ' ', '') LIKE '%'+@searchString+'%' OR Replace(ltitle, ' ', '') LIKE '%'+@searchString+'%' )";
            }
            sql += @"
                                 ),
                            materialPerPage
                              AS ( SELECT   ROW_NUMBER() OVER ( ORDER BY " + orderType + @" DESC) AS erank ,
                                            (SELECT COUNT(*) AS totalData FROM materialPage)AS totalData,
                                            *
                                   FROM     materialPage WITH ( NOLOCK )
                                 )
                        SELECT  *
                        FROM    materialPerPage WITH ( NOLOCK )
                        WHERE   erank >= @PAGEStart
                                AND erank <= @PAGEEnd
            ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@sessionType", SqlDbType.VarChar);
            cmd.Parameters.Add("@attend_level", SqlDbType.VarChar);
            cmd.Parameters.Add("@ltopic", SqlDbType.VarChar);
            cmd.Parameters.Add("@searchString", SqlDbType.VarChar);
            cmd.Parameters.Add("@PAGE", SqlDbType.Int);
            cmd.Parameters.Add("@PAGEEnd", SqlDbType.Int);
            cmd.Parameters.Add("@PAGEStart", SqlDbType.Int);

            cmd.Parameters["@sessionType"].Value = sessionType;
            cmd.Parameters["@attend_level"].Value = materialLevel;
            cmd.Parameters["@ltopic"].Value = interestType;
            cmd.Parameters["@searchString"].Value = searchString;
            cmd.Parameters["@PAGE"].Value = intPage;
            cmd.Parameters["@PAGEEnd"].Value = intPage * CommData.videoPerPage;
            cmd.Parameters["@PAGEStart"].Value = intPage * CommData.videoPerPage - (CommData.videoPerPage - 1);
            rs = cmd.ExecuteReader();
            while (rs.Read())
            {
                videoInformation.Add(GetVideoInfoFromDataReader(rs));
                count = Convert.ToInt32(rs["totalData"].ToString());
            }
            setVideoLength(Convert.ToInt32(count));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return FilterBySearchInfo(videoInformation);
    }
    /*private bool IsVideoUpdate()
    {
		return false;
        int videoCount = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;

        try
        {
            string sql = " SELECT Count(session_type) AS videoCount FROM SessionRecord_selected_by_material(nolock) ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;

            while (rs.Read())
            {
                videoCount = Convert.ToInt32(rs["videoCount"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();

        }
        if (videoCount != intSelectLength)
        {
            intSelectLength = videoCount;
			return true;
        }
        return false;   
    }*/

    public double GetUserFreePoint(string clinet_sn)
    {

        double freePoint = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SUM(used_video_point) AS last_free_point FROM client_free_video_point WHERE (client_sn =" + clinet_sn + ") ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return 0 ;
            while (rs.Read())
            {
                if(rs["last_free_point"]!=DBNull.Value && !rs["last_free_point"].ToString().Equals(""))
                      freePoint += Convert.ToDouble(rs["last_free_point"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return freePoint;
    }

    private VideoInfo GetVideoInfoFromDataReader(SqlDataReader rs)
    {
        VideoInfo videoInfo = new VideoInfo();
        try
        {
            videoInfo.Consultant = rs["consultant"].ToString();
            videoInfo.Course = rs["course"].ToString();
            videoInfo.File_FullName = rs["File_Full_name"].ToString();
            videoInfo.LVList = new ArrayList(rs["attend_level"].ToString().Split(','));
            videoInfo.NameOfVideo = rs["ltitle"].ToString();
            videoInfo.SessionRating = Convert.ToDouble(rs["Rating"].ToString());
            videoInfo.SessionSN = rs["session_sn"].ToString();
            videoInfo.SessionStudent = new ArrayList(rs["client_sn_list"].ToString().Split(','));
            videoInfo.SessionType = rs["session_type"].ToString();
            videoInfo.TopicList = new ArrayList(rs["ltopic"].ToString().Split(','));
            videoInfo.File_SN = rs["file_sn"].ToString();
            videoInfo.File_Path = rs["file_path"].ToString();
            videoInfo.File_Date = DateTime.Parse(rs["file_date"].ToString());
            if (rs["fg"] != DBNull.Value)
                videoInfo.Focus_Grammar = Convert.ToInt32(rs["fg"].ToString());
            if (rs["fl"] != DBNull.Value)
                videoInfo.Focus_Listening = Convert.ToInt32(rs["fl"].ToString());
            if (rs["fp"] != DBNull.Value)
                videoInfo.Focus_Pronuciation = Convert.ToInt32(rs["fp"].ToString());
            if (rs["fr"] != DBNull.Value)
                videoInfo.Focus_Reading = Convert.ToInt32(rs["fr"].ToString());
            if (rs["fs"] != DBNull.Value)
                videoInfo.Focus_Speaking = Convert.ToInt32(rs["fs"].ToString());
            if (rs["fv"] != DBNull.Value)
                videoInfo.Focus_Vocabulary = Convert.ToInt32(rs["fv"].ToString());
            if (rs["description"] != DBNull.Value && !rs["description"].ToString().Equals(""))
                videoInfo.Description = rs["description"].ToString().Replace("<br>", "\n");
      
            if (videoInfo.SessionStudent.IndexOf(clientSn) != -1)
                videoInfo.IsOwnSession = true;

            videoInfo.Mtl_RatingInfo = GetRatingInfo(videoInfo.Course, videoInfo.SessionSN);
            videoInfo.CourseCount = rs["CourseCount"].ToString();
        }
        catch
        {
            videoInfo = null;
        }
        finally
        {
            
        }
        return videoInfo;
    }

    /*public VideoInfo[] Top15VideoList()
    {
   
        int max_topCount = CommData.videoPerPage;
        int topCount = 0;
        ArrayList top15List = new ArrayList();
        ReadFromSelectedTable();
        try
        {
            VideoInfo[] searchedSpecialSessionVideoInfo = FilterBySearchInfo(CommData.GetallSpecialSessionVideoInfo(), true);
            VideoInfo[] searchedRgVideoInfo = FilterBySearchInfo((ArrayList)CommData.GetallRegularlSessionVideoInfo(), true);
            VideoInfo[] searchedMySessionVideoInfo = MySessionVideoList();
            topCount = 0;
            foreach (VideoInfo aVideo in searchedSpecialSessionVideoInfo)
            {

                topCount++;
                top15List.Add(aVideo);

                if (topCount == max_topCount)
                    break;
            }
            topCount = 0;
            foreach (VideoInfo aVideo in searchedRgVideoInfo)
            {
                topCount++;
                top15List.Add(aVideo);
                if (topCount == max_topCount)
                    break;
            }
            topCount = 0;
            foreach (VideoInfo aVideo in searchedMySessionVideoInfo)
            {
                topCount++;
                top15List.Add(aVideo);
                if (topCount == max_topCount)
                    break;
            }
            return (VideoInfo[])top15List.ToArray(typeof(VideoInfo));
        }
        finally
        {
            top15List.Clear();
            top15List = null;
        }
    }*/

    private bool IsClientView(string course, string session_sn)
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SessionRecord_fileinfo.Session_Record_Number, sessionrecord_view_list.rating, sessionrecord_view_list.Course " +
                               " FROM sessionrecord_view_list INNER JOIN " +
                               " SessionRecord_fileinfo ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.SN " +
                               " WHERE (SessionRecord_fileinfo.Session_Record_Number =@session_sn) AND (sessionrecord_view_list.Course =@course) AND (sessionrecord_view_list.client_sn =@client_sn) ";

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters["@session_sn"].Value = session_sn;
            cmd.Parameters["@course"].Value = course;
            cmd.Parameters["@client_sn"].Value = clientSn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    private int GetViewCount(string course, string session_sn)
    {
        int viewCount = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT COUNT(SessionRecord_fileinfo.Session_Record_Number) AS Session_Record_Number/*, sessionrecord_view_list.rating, sessionrecord_view_list.Course*/ " +
                               " FROM sessionrecord_view_list INNER JOIN " +
                               " SessionRecord_fileinfo ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.SN " +
                               " WHERE (SessionRecord_fileinfo.Session_Record_Number =@session_sn) AND (sessionrecord_view_list.Course =@course) ";

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters["@session_sn"].Value = session_sn;
            cmd.Parameters["@course"].Value = course;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return -1;
            while (rs.Read())
            {
                viewCount = Convert.ToInt32(rs["Session_Record_Number"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return viewCount;
    }

    private RatingInfo GetRatingInfo(string course, string session_sn)
    {
        RatingInfo ratingInfo = new RatingInfo();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = @" 
                SELECT SessionRecord_fileinfo.Session_Record_Number, sessionrecord_view_list.rating, sessionrecord_view_list.Course,sessionrecord_view_list.rating_date,sessionrecord_view_list.client_sn
                FROM sessionrecord_view_list WITH(NOLOCK) INNER JOIN 
                SessionRecord_fileinfo WITH(NOLOCK) ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.SN 
                WHERE (SessionRecord_fileinfo.Session_Record_Number =@session_sn) AND (sessionrecord_view_list.Course =@course)
                AND client_sn = @client_sn
            ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters["@session_sn"].Value = session_sn;
            cmd.Parameters["@course"].Value = course;
            cmd.Parameters["@client_sn"].Value = ClientSn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return ratingInfo;
            while (rs.Read())
            {
                if (rs["rating_date"] != DBNull.Value && rs["rating"] != DBNull.Value)
                {
                    ratingInfo.UsersRatingList.Add(Convert.ToDouble(rs["rating"].ToString()));
                    ratingInfo.UsersRatingDate.Add(DateTime.Parse(rs["rating_date"].ToString()));
                    ratingInfo.UserRating = Convert.ToDouble(rs["rating"].ToString());;
                }
               //if(clientSn.Equals(rs["client_sn"].ToString()))
                   ratingInfo.IsView =true;
                //ratingInfo.ViewCount++;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }

        return ratingInfo;
    }

    private double GetUserRating(string session_sn)
    {
        double userRating = -1;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SessionRecord_fileinfo.Session_Record_Number, sessionrecord_view_list.rating " +     
                               " FROM sessionrecord_view_list INNER JOIN " +
                               " SessionRecord_fileinfo ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.SN " +
                               " WHERE (sessionrecord_view_list.client_sn = @client_sn) AND (SessionRecord_fileinfo.Session_Record_Number =@session_sn) ";

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters["@client_sn"].Value = clientSn;
            cmd.Parameters["@session_sn"].Value = session_sn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return -1;
            if(rs.Read())
            {
                userRating = Convert.ToDouble(rs["rating"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return userRating;
    }

    private double GetAverageUsersRating(string course, string session_sn)
    {
        double viewCount = 0;
        double totalUserRating = 0;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT SessionRecord_fileinfo.Session_Record_Number, sessionrecord_view_list.rating, sessionrecord_view_list.Course " +
                               " FROM sessionrecord_view_list INNER JOIN " +
                               " SessionRecord_fileinfo ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.SN " +
                               " WHERE (SessionRecord_fileinfo.Session_Record_Number =@session_sn) AND (sessionrecord_view_list.Course =@course) ";

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@course", SqlDbType.VarChar);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters["@course"].Value = course;
            cmd.Parameters["@session_sn"].Value = session_sn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return -1;
            while (rs.Read())
            {
                viewCount++;
                totalUserRating += Convert.ToDouble(rs["rating"].ToString());
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return totalUserRating / viewCount;
    }

    /*public VideoInfo[] SelectedSpecialSessionVideoList()//要改
    {
        ReadFromSelectedTable();
        return FilterBySearchInfo(CommData.GetallSpecialSessionVideoInfo(),true);
    }*/

    public bool AddUserRating(VideoInfo videoInfo,int rating)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " UPDATE sessionrecord_view_list SET rating = @rating,rating_date = @rating_date"+
                              " WHERE course = @course AND client_sn =@client_sn";

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@course", SqlDbType.VarChar);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@rating", SqlDbType.Int);
            cmd.Parameters.Add("@rating_date", SqlDbType.DateTime);
          
            cmd.Parameters["@course"].Value = videoInfo.Course;
            cmd.Parameters["@client_sn"].Value = clientSn;
            cmd.Parameters["@rating"].Value = rating;
            cmd.Parameters["@rating_date"].Value = DateTime.Now;

           
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public VideoInfo GetVideoInfoBySessionSN(string sessionSN)
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        VideoInfo videoInfo = new VideoInfo();
        try
        {
            string sql = "SELECT * FROM vSessionVideosVIPABC WITH(NOLOCK) WHERE session_sn=@session_sn ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters["@session_sn"].Value = sessionSN;
            rs = cmd.ExecuteReader();
            /*if (rs == null)
                return null ;*/
            while (rs.Read())
            {
                videoInfo = GetVideoInfoFromDataReader(rs);
                if (videoInfo != null)
                {//此方法暫時無法更動，除非原來就傳入VideoInfo類別，再SQL，因要花時間去查哪裡有用這個function，因此暫時使用原本的方案。
                    foreach (string aStudentSn in videoInfo.SessionStudent)
                    {
                        if (aStudentSn.Equals(clientSn))
                        {
                            videoInfo.IsOwnSession = true;
                            break;
                        }
                    }
                    return videoInfo;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return videoInfo;
    }

    /*public VideoInfo[] AllRegSelectedVideoList()
    {
        ReadFromSelectedTable();
        ArrayList resultList = new ArrayList();
        try
        {
            foreach (VideoInfo aVideoInfo in CommData.GetallRegularlSessionVideoInfo())
            {
                resultList.Add(aVideoInfo.CloneVideoInfo());
            }
            return FilterBySearchInfo(resultList, true);
        }
        finally
        {
            resultList.Clear();
            resultList = null;
        }
    }

	public VideoInfo[] AllFSSelectedVideoList()
	{
		ReadFromSelectedTable();
		ArrayList resultList = new ArrayList();
		try
		{
			foreach (VideoInfo aVideoInfo in CommData.GetallFamousSessionVideoInfo())
			{
				resultList.Add(aVideoInfo.CloneVideoInfo());
			}
			return FilterBySearchInfo(resultList, true);
		}
		finally
		{
			resultList.Clear();
			resultList = null;
		}
	}

    public VideoInfo[] AllSelectedVideoList()
    {
        ReadFromSelectedTable();
        ArrayList resultList = new ArrayList();
        try
        {
            foreach (VideoInfo aVideoInfo in CommData.GetallRegularlSessionVideoInfo())
            {
                resultList.Add(aVideoInfo.CloneVideoInfo());
            }
            foreach (VideoInfo aVideoInfo in CommData.GetallSpecialSessionVideoInfo())
            {
                resultList.Add(aVideoInfo.CloneVideoInfo());
            }
			foreach (VideoInfo aVideoInfo in CommData.GetallFamousSessionVideoInfo())
			{
				resultList.Add(aVideoInfo.CloneVideoInfo());
			}
            return FilterBySearchInfo(resultList, true);//把以上的LIST丟進去
        }
        finally
        {
            resultList.Clear();
            resultList = null;
        }
    }*/

   /*private VideoInfo[] FilterBySearchInfo1(ArrayList orgVideoList,bool sortResult)//
    {
        ArrayList tempList = (ArrayList)orgVideoList.Clone();
		 ArrayList resultList =  new ArrayList();
        try
        {
            if (!clientSearchInfo.VideoInterestType.Equals(""))
            {
                foreach (VideoInfo aVideoInfo in orgVideoList)
                {
                    bool isMatchInterest = false;
                    foreach (string aInterest in aVideoInfo.TopicList)
                    {

                        if (aInterest.Equals(clientSearchInfo.VideoInterestType))
                        {
                            isMatchInterest = true;
                        }
                    }
                    if (!isMatchInterest)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (!clientSearchInfo.VideoLevel.ToString().Equals("0"))
            {
                foreach (VideoInfo aVideoInfo in orgVideoList)
                {
                    bool isMatchLevel = false;
                    foreach (string aLev in aVideoInfo.LVList)
                    {

                        if (aLev.Equals(clientSearchInfo.VideoLevel.ToString()))
                        {
                            aVideoInfo.SessionLevel = Convert.ToInt32(aLev);
                            isMatchLevel = true;

                        }
                    }
                    if (!isMatchLevel)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (clientSearchInfo.Fous_Type != null && !clientSearchInfo.Fous_Type.Equals(""))
            {
                foreach (VideoInfo aVideoInfo in orgVideoList)
                {
                    bool isMatchFocus = false;
                    foreach (string aFocus in aVideoInfo.Focus_List)
                    {
                        if (aFocus.Equals(clientSearchInfo.Fous_Type))
                        {
                            isMatchFocus = true;
                        }
                    }
                    if (!isMatchFocus)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (!clientSearchInfo.SearchPatthen.Equals(""))
            {
                foreach (VideoInfo aVideoInfo in orgVideoList)
                {
                    if (aVideoInfo.NameOfVideo.Trim().ToLower().IndexOf(clientSearchInfo.SearchPatthen.Trim().ToLower()) == -1 && aVideoInfo.Consultant.Trim().ToLower().IndexOf(clientSearchInfo.SearchPatthen.Trim().ToLower()) == -1)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (sortResult)
            {
                if (clientSearchInfo.RankType.Equals("highRating"))
                {
                    tempList.Sort(new myHighRatingClass());
                }
                else if (clientSearchInfo.RankType.Equals("highView"))
                {
                    tempList.Sort(new myHighViewClass());
                }
                else
                {
                    tempList.Sort(new myNewestClass());
                }
            }
            foreach (VideoInfo aVideo in tempList)
            { 
                foreach (string aStudentSn in aVideo.SessionStudent)
                {
                    if (aStudentSn.Equals(clientSn))
                    {
                        aVideo.IsOwnSession = true;
                        break;
                    }

                }
            }
            foreach (VideoInfo aVideo in tempList)
            {
                if (this.IsStudySameSeesionSn(aVideo))
                    aVideo.IsClientUseThisSessionSn = true;
            }
            foreach (VideoInfo aVideo in tempList)
            {
                if (this.IsStudySameMaterial(aVideo))
                    aVideo.IsClientUseThisMaterial = true;
            }
            foreach (VideoInfo aVideo in tempList)
            {
                if (this.IsStudySameRecordFile(aVideo))
                    aVideo.IsClientUseThisRecordFile = true;
            }
            return (VideoInfo[])tempList.ToArray(typeof(VideoInfo));
        }
        finally
        {
            tempList.Clear();
            tempList = null;
        }
    }*/

    private List<VideoInfo> FilterBySearchInfo(List<VideoInfo> orgVideoList)
    {//乾...
        //暫時無視
        /*ArrayList tempList = (ArrayList)orgVideoList.Clone();
        try {
            if (!clientSearchInfo.VideoInterestType.Equals("")) {
                foreach (VideoInfo aVideoInfo in orgVideoList) {
                    bool isMatchInterest = false;
                    foreach (string aInterest in aVideoInfo.TopicList) {

                        if (aInterest.Equals(clientSearchInfo.VideoInterestType)) {
                            isMatchInterest = true;
                            break;
                        }
                    }
                    if (!isMatchInterest)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (!clientSearchInfo.VideoLevel.ToString().Equals("0")) {
                foreach (VideoInfo aVideoInfo in orgVideoList) {
                    bool isMatchLevel = false;
                    foreach (string aLev in aVideoInfo.LVList) 
                    {
                        if (aLev.Equals(clientSearchInfo.VideoLevel.ToString())) {
                            aVideoInfo.SessionLevel = Convert.ToInt32(aLev);
                            isMatchLevel = true;
                            break;
                        }
                    }
                    if (!isMatchLevel)
                        tempList.Remove(aVideoInfo);
                }
            }*/
            /*if (clientSearchInfo.Fous_Type != null && !clientSearchInfo.Fous_Type.Equals("")) {
                foreach (VideoInfo aVideoInfo in orgVideoList) {
                    bool isMatchFocus = false;
                    foreach (string aFocus in aVideoInfo.Focus_List) {
                        if (aFocus.Equals(clientSearchInfo.Fous_Type)) {
                            isMatchFocus = true;
                            break;
                        }
                    }
                    if (!isMatchFocus)
                        tempList.Remove(aVideoInfo);
                }
            }*/
           //暫時無視
           /*
            if (!clientSearchInfo.SearchPatthen.Equals("")) {
                foreach (VideoInfo aVideoInfo in orgVideoList) {
                    if (aVideoInfo.NameOfVideo.Trim().ToLower().IndexOf(clientSearchInfo.SearchPatthen.Trim().ToLower()) == -1 && aVideoInfo.Consultant.Trim().ToLower().IndexOf(clientSearchInfo.SearchPatthen.Trim().ToLower()) == -1)
                        tempList.Remove(aVideoInfo);
                }
            }
            if (sortResult) {
                if (clientSearchInfo.RankType.Equals("highRating")) {
                    tempList.Sort(new myHighRatingClass());
                } else if (clientSearchInfo.RankType.Equals("highView")) {
                    tempList.Sort(new myHighViewClass());
                } else {
                    tempList.Sort(new myNewestClass());
                }
            }*/

            List<ClientLessonData> attendedSessions = LessonDataHelper.GetAttendSessions(ClientSn);
            List<string> recordFiles = LessonDataHelper.GetRecordFile(ClientSn);

            foreach (VideoInfo aVideo in orgVideoList) {
                foreach (string aStudentSn in aVideo.SessionStudent) {
                    if (aStudentSn.Equals(clientSn)) {
                        aVideo.IsOwnSession = true;
                        break;
                    }
                }
                foreach (ClientLessonData data in attendedSessions) {
                    if (aVideo.SessionSN.Equals(data.SessionSN)) {
                        aVideo.IsClientUseThisSessionSn = true;
                        break;
                    }
                }
                foreach (ClientLessonData data in attendedSessions) {
                    if (aVideo.Course.Equals(data.MaterialSN)) {
                        aVideo.IsClientUseThisMaterial = true;
                        break;
                    }
                }
                foreach (string file in recordFiles) {
                    if (aVideo.Course.Equals(file)) {
                        aVideo.IsClientUseThisRecordFile = true;
                        break;
                    }
                }
            }
            return orgVideoList;
    }   

    /*public VideoInfo[] MySessionVideoList()
    {
        ArrayList mySessionVideoList = new ArrayList();
        try
        {
         ReadFromSelectedTable();
            if (mySessionVideoList.Count > 0)
                return (VideoInfo[])mySessionVideoList.ToArray(typeof(VideoInfo));
            foreach (VideoInfo aVideo in CommData.GetallPackageSessionVideoInfo())
            {
                if (aVideo.SessionStudent.IndexOf(clientSn) != -1)
                {
                    VideoInfo mySessionaVideo = aVideo.CloneVideoInfo();
                    mySessionaVideo.IsOwnSession = true;
                    mySessionVideoList.Add(mySessionaVideo);
                }
            }           
            return (VideoInfo[])mySessionVideoList.ToArray(typeof(VideoInfo));            
        }
        finally
        {
            //tempList.Clear();
            //tempList = null;
            mySessionVideoList.Clear();
            mySessionVideoList = null;
        }
    }*/

    public bool IsUnlimitDownloadTime(VideoInfo videoInfo, string client_sn)
    {

        bool isUunLimitdownload  =   IsFromPaidList(videoInfo, client_sn);
        if (isUunLimitdownload)
        {
            return isUunLimitdownload;
        }
        return IsFromFreeList(videoInfo, client_sn);

    }

    private bool IsFromFreeList(VideoInfo videoInfo, string client_sn)  //扣贈送堂數
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT DateDiff(n, view_datetime, GETDATE()) AS record_connect_time, view_datetime FROM sessionrecord_view_used_free_video_point_list WHERE (client_sn = '" + client_sn + "') AND (Course = '" + videoInfo.Course + "') ORDER BY SN DESC ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            while (rs.Read())
            {
                if (Convert.ToInt32(rs["record_connect_time"].ToString()) <= CommData.unlimitDownloadSec)
                    return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    private  bool IsFromPaidList(VideoInfo videoInfo, string client_sn)     //扣一般堂數
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT DateDiff(n, view_datetime, GETDATE()) AS record_connect_time, view_datetime FROM sessionrecord_view_list WHERE (client_sn = '" + client_sn + "') AND (Course = '" + videoInfo.Course + "') and valid =1 ORDER BY SN DESC ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if ( rs == null )
                return false;
            while (rs.Read())
            {
                if (Convert.ToInt32(rs["record_connect_time"].ToString()) <= CommData.unlimitDownloadSec)
                    return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    public bool AddtoFreeViewList(VideoInfo videoInfo, string client_sn, double cost, string client_ip)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " INSERT INTO sessionrecord_view_free_list(client_sn,view_cost,view_session_recordfiles_sn,view_datetime,view_ip,Course,view_state)" +
                              " VALUES(@client_sn,@view_cost,@view_session_recordfiles_sn,@view_datetime,@view_ip,@Course,@view_state)";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_cost", SqlDbType.Float);
            cmd.Parameters.Add("@view_session_recordfiles_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_datetime", SqlDbType.DateTime);
            cmd.Parameters.Add("@view_ip", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@view_state", SqlDbType.VarChar);
            cmd.Parameters["@client_sn"].Value = client_sn;
            cmd.Parameters["@view_cost"].Value = cost;
            cmd.Parameters["@view_session_recordfiles_sn"].Value = videoInfo.File_SN;
            cmd.Parameters["@view_datetime"].Value = DateTime.Now;
            cmd.Parameters["@view_ip"].Value = client_ip;
            cmd.Parameters["@Course"].Value = videoInfo.Course;
            cmd.Parameters["@view_state"].Value = "1";
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public bool AddtoFreePointViewList(VideoInfo videoInfo, string client_sn,double cost,string client_ip)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " INSERT INTO sessionrecord_view_used_free_video_point_list(client_sn,view_cost,view_session_recordfiles_sn,view_datetime,view_ip,Course,view_state)" +
                               " VALUES(@client_sn,@view_cost,@view_session_recordfiles_sn,@view_datetime,@view_ip,@Course,@view_state)";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_cost", SqlDbType.Float);
            cmd.Parameters.Add("@view_session_recordfiles_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_datetime", SqlDbType.DateTime);
            cmd.Parameters.Add("@view_ip", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@view_state", SqlDbType.VarChar);
            cmd.Parameters["@client_sn"].Value = client_sn;
            cmd.Parameters["@view_cost"].Value = cost;
            cmd.Parameters["@view_session_recordfiles_sn"].Value =videoInfo.File_SN;
            cmd.Parameters["@view_datetime"].Value = DateTime.Now;
            cmd.Parameters["@view_ip"].Value = client_ip;
            cmd.Parameters["@Course"].Value = videoInfo.Course;
            cmd.Parameters["@view_state"].Value = "1";
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public bool AddtoCostViewList(VideoInfo videoInfo, string client_sn, double cost, string client_ip, string vds_seq,string vds_video_type)
    {
        bool isok = false;
        string content = "error";
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            //'======= 工單:BD09050122、BD09050603 特殊扣點觀看錄影檔設定 加入vds_video_type vds_seq raychien 20090511 START =========
            string sql = "";
            string intViewListSn = "";
            string urlStr = "";
            if (String.IsNullOrEmpty(vds_seq) == true)
            {
                sql = " INSERT INTO sessionrecord_view_list(client_sn,view_cost,view_session_recordfiles_sn,view_datetime,view_ip,Course,view_state, vds_video_type)" +
                        " VALUES(@client_sn,@view_cost,@view_session_recordfiles_sn,@view_datetime,@view_ip,@Course,@view_state, @vds_video_type) SELECT SCOPE_IDENTITY() AS sn ";
            }
            else
            {
                sql = " INSERT INTO sessionrecord_view_list(client_sn,view_cost,view_session_recordfiles_sn,view_datetime,view_ip,Course,view_state, vds_video_type, vds_seq)" +
                        " VALUES(@client_sn,@view_cost,@view_session_recordfiles_sn,@view_datetime,@view_ip,@Course,@view_state, @vds_video_type, @vds_seq) SELECT SCOPE_IDENTITY() AS sn ";
            }
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);

            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_cost", SqlDbType.Float);
            cmd.Parameters.Add("@view_session_recordfiles_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_datetime", SqlDbType.DateTime);
            cmd.Parameters.Add("@view_ip", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@view_state", SqlDbType.VarChar);

            cmd.Parameters.Add("@vds_video_type", SqlDbType.VarChar);
            if (String.IsNullOrEmpty(vds_seq) == false)
            {
                cmd.Parameters.Add("@vds_seq", SqlDbType.Char);
            }

            cmd.Parameters["@client_sn"].Value = client_sn;
            cmd.Parameters["@session_sn"].Value = videoInfo.SessionSN;
            cmd.Parameters["@view_cost"].Value = cost;
            cmd.Parameters["@view_session_recordfiles_sn"].Value = videoInfo.File_SN;
            cmd.Parameters["@view_datetime"].Value = DateTime.Now;
            cmd.Parameters["@view_ip"].Value = client_ip;
            cmd.Parameters["@Course"].Value = videoInfo.Course;
            cmd.Parameters["@view_state"].Value = "1";
            cmd.Parameters["@vds_video_type"].Value = vds_video_type;
            if (String.IsNullOrEmpty(vds_seq) == false)
            {
                cmd.Parameters["@vds_seq"].Value = vds_seq;
            }
            //cmd.ExecuteNonQuery();
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            while (rs.Read())
            {
                intViewListSn = Convert.ToInt32(rs["sn"]).ToString();
            }
            //20121127 阿捨新增 傳送html
            urlStr = "http://www.vipabc.com/program/member/reservation_class/metainVideoRecordPoint.asp?client_sn=" + Convert.ToString(client_sn) + "&int_session=" + Convert.ToString(cost) + "&sessionrecord_view_list=" + intViewListSn;
            content = GetHtml(urlStr).ToString();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public bool AddToSessionrecord_view_used_free_video_point_list(VideoInfo videoInfo, string client_sn, double cost, string client_ip)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " INSERT INTO sessionrecord_view_used_free_video_point_list(client_sn,view_cost,view_session_recordfiles_sn,view_datetime,view_ip,Course,view_state)" +
                               " VALUES(@client_sn,@view_cost,@view_session_recordfiles_sn,@File_Full_name,@view_datetime,@Course,@view_state)";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_cost", SqlDbType.Float);
            cmd.Parameters.Add("@view_session_recordfiles_sn", SqlDbType.VarChar);
            cmd.Parameters.Add("@view_datetime", SqlDbType.DateTime);
            cmd.Parameters.Add("@view_ip", SqlDbType.VarChar);
            cmd.Parameters.Add("@Course", SqlDbType.Int);
            cmd.Parameters.Add("@view_state", SqlDbType.VarChar);
            cmd.Parameters["@client_sn"].Value = client_sn;
            cmd.Parameters["@view_cost"].Value = cost;
            cmd.Parameters["@view_session_recordfiles_sn"].Value = videoInfo.File_SN;
            cmd.Parameters["@view_datetime"].Value = DateTime.Now;
            cmd.Parameters["@view_ip"].Value = client_ip;
            cmd.Parameters["@Course"].Value = videoInfo.Course;
            cmd.Parameters["@view_state"].Value = "1";
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    public bool IsStudySameSeesionSn(VideoInfo videoInfom)
    {
        bool isUsed = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT * FROM client_attend_list WHERE (client_sn =@client_sn) AND session_sn=@session_sn AND valid=1";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters.Add("@session_sn", SqlDbType.VarChar);
            cmd.Parameters["@client_sn"].Value = clientSn;
            cmd.Parameters["@session_sn"].Value = videoInfom.SessionSN;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                isUsed = true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isUsed;
    }

    public bool IsStudySameMaterial(VideoInfo videoInfom)
    {
        bool isUsed = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT attend_mtl_1 FROM client_attend_list WHERE (client_sn =@client_sn) AND (attend_mtl_1 =@attend_mtl_1) AND valid=1";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters.Add("@attend_mtl_1", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = clientSn;
            cmd.Parameters["@attend_mtl_1"].Value = videoInfom.Course;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                isUsed = true;
            }
        }
        catch (Exception ex)
        {

            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isUsed;
    }

    public bool IsStudySameRecordFile(VideoInfo videoInfom)
    {
        bool isUsed = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT course FROM sessionrecord_view_list WHERE (client_sn =@client_sn) AND (course =@course) AND valid= 1           ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters.Add("@course", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = clientSn;
            cmd.Parameters["@course"].Value = videoInfom.Course;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            if (rs.Read())
            {
                isUsed = true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isUsed;
    }

    public bool DecreaseFreePoint(VideoInfo videoInfo, string client_sn, float cost)
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT used_video_point,sn FROM client_free_video_point WHERE (client_sn = "+client_sn +") AND (used_video_point > 0) ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return false;
            while (rs.Read())
            {
                float free_point = (float)Convert.ToDouble(rs["used_video_point"].ToString());
                if (free_point >= cost)
                {
                    UpdateClientFreePoint(rs["sn"].ToString(), free_point - cost);
                    break;
                }
                else
                {
                    UpdateClientFreePoint(rs["sn"].ToString(), 0);
					cost = cost - free_point;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return false;
    }

    public string GetClientLevel()
    {
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT nlevel FROM client_basic WHERE (sn = " + clientSn + ")";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return "6";
            while (rs.Read())
            {
                return rs["nlevel"].ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return "6";
    }

    public string GetClientInterest()
    {
        string[] interests = null;
        ArrayList interestList = new ArrayList();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = "SELECT interest FROM client_basic WHERE (sn = " + clientSn + ")";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            if (rs == null)
                return "1";
            while (rs.Read())
            {
                string interest =rs["interest"].ToString();
                if (interest.IndexOf(",")!=-1)
                {
                    interests = interest.Split(',');
                }
                if (interest.IndexOf("-")!=-1)
                {
                    interest = interest.Replace("--", "-");
                    interests = interest.Split('-');
                }
            }
            foreach (string aInterest in interests)
            {
                if (!aInterest.Equals(""))
                    interestList.Add(aInterest);
            }
            if (interestList.Count > 0)
                return interestList[0].ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
            interestList.Clear();
            interests = null;
        }
        return "1";
    }

    private bool UpdateClientFreePoint(string sn, float newPoint)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            string sql = " UPDATE client_free_video_point SET used_video_point = @used_video_point" +
                              " WHERE sn = @sn";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@used_video_point", SqlDbType.Float);
            cmd.Parameters.Add("@sn", SqlDbType.Int);
            cmd.Parameters["@used_video_point"].Value = newPoint;
            cmd.Parameters["@sn"].Value = Convert.ToInt32(sn);
            cmd.ExecuteNonQuery();
            isok = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return isok;
    }

    //工單:BD09050122、BD09050603 特殊扣點觀看錄影檔設定 raychien 20090511
    public float GetPointsFromDownloadVideoRule(int ClientSn, int VideoType, string VideoDownloadSettingSeq, string OtherRule)
    //'函式描述: 根據為客戶設定特殊扣點觀看錄影檔原則 傳回需扣的點數(Float)
    //'Input: ClientSn：客戶編號
    //'          VideoType：錄影檔類型 [video_download_setting.vds_video_type]
    //'          VideoDownloadSettingSeq：[video_download_setting.vds_seq]
    //'          OtherRule：額外的規則(只有定義，尚未實做，可傳入空值)
    //' Ouput: Success: 回傳 Points As String
    //' Failed: 回傳 -1，表示SQL執行失敗
    //'             回傳 -2，表示沒有為此客戶設定任何扣點規則
    //'             回傳 -3，表示傳入的ClientSn或VideoType是空值
    //'             回傳 -4，表示為此客戶設定的規則的下載次數已達上限(包括不拘)
    //'             回傳 -5，表示為此客戶設定的規則的下載堂數已達上限(包括不拘)
    //'             回傳 -9，例外錯誤
    //' 20130206 阿捨新增 實做上限堂數判斷 並新增不拘類型(9)之總錄影檔判斷
    {
        bool bo_error_happened;
        //int vds_allow_course_count;
        int client_sn = 0, video_type = 0, vds_allow_download_count = 0, used_count = 0, vds_video_type = 0, intSessionTypeSn = 0;
        string vds_seq = "", str_sql_statement = "", str_condition = "", str_sql_statement1 = "", str_sql_statement2 = "",  str_condition1 = "";
        float vds_subtract_points = 0, vds_allow_course_count = 0, ouput_value = 0, used_cost = 0, gift_total_count = 0;
        DateTime vds_valid_sdate = new DateTime();
        DateTime vds_valid_edate = new DateTime();

        try
        {
            //'是否有錯誤發生
            bo_error_happened = false;

            //'回傳值
            ouput_value = 0;

            client_sn = ClientSn;
            video_type = VideoType;
            vds_seq = VideoDownloadSettingSeq;

            if ( String.IsNullOrEmpty(Convert.ToString(client_sn)) == true || String.IsNullOrEmpty(Convert.ToString(video_type)) == true )
            {
                bo_error_happened = true;
                ouput_value = -3;
            }

            if ( bo_error_happened == false )
            {
                ArrayList paraList = new ArrayList();
                paraList.Add(CreateSqlParameter("@client_sn", client_sn));
                //paraList.Add(CreateSqlParameter("@vds_seq", vds_seq));
                paraList.Add(CreateSqlParameter("@video_type", video_type));

                str_condition = " AND ( vds_video_type = @video_type ) ";

                //'取出為客戶設定特殊扣點觀看錄影檔原則 最新的一筆
                str_sql_statement = " SELECT TOP 1 vds_valid_sdate, vds_valid_edate, vds_subtract_points, ISNULL(vds_allow_download_count, 0) AS vds_allow_download_count, ISNULL(vds_allow_course_count, 0) AS vds_allow_course_count, vds_video_type ";
                str_sql_statement += " FROM video_download_setting WITH (NOLOCK) ";
                str_sql_statement += " WHERE ( vds_valid = 1 ) ";
                str_sql_statement += " AND ( CONVERT(VARCHAR, vds_valid_sdate, 111) <= CONVERT(VARCHAR, GETDATE(), 111) ) ";
                str_sql_statement += " AND ( CONVERT(VARCHAR, vds_valid_edate, 111) >= CONVERT(VARCHAR, GETDATE(), 111) ) ";
                str_sql_statement += " AND ( clb_sn = @client_sn ) ";
                str_sql_statement += str_condition;
                str_sql_statement += " ORDER BY vds_create_datetime DESC ";

                DataTable user_SessionTable = ExcuteSql(str_sql_statement, paraList);
                DataTable user_SessionTable1;
                DataTable user_SessionTable2;

                // 例外錯誤
                if ( user_SessionTable == null )
                {
                    return -9;
                }
                if ( user_SessionTable.Rows.Count > 0 )
                {
                    foreach ( DataRow aRow in user_SessionTable.Rows )
                    {
                        vds_valid_sdate = Convert.ToDateTime(aRow["vds_valid_sdate"].ToString()); //'開始日期
                        vds_valid_edate = Convert.ToDateTime(aRow["vds_valid_edate"].ToString()); //'結束日期
                        vds_allow_course_count = (float)Convert.ToDouble(aRow["vds_allow_course_count"]);  //'限制的堂數
                        vds_allow_download_count = (int)Convert.ToInt32(aRow["vds_allow_download_count"]); //'限制下載的次數
                        vds_subtract_points = (float)Convert.ToDouble(aRow["vds_subtract_points"]); //'要扣的點數
                        vds_video_type = (int)Convert.ToInt32(aRow["vds_video_type"]); //'錄影檔類型

                        //'取得已下載錄影檔的次數及堂數
                        paraList.Clear();
                        paraList.Add(CreateSqlParameter("@client_sn2", client_sn));
                        paraList.Add(CreateSqlParameter("@vds_valid_sdate2", vds_valid_sdate.ToShortDateString()));
                        paraList.Add(CreateSqlParameter("@vds_valid_edate2", vds_valid_edate.ToShortDateString()));

                        //' 次數 或 堂數 無法共存 擇一判斷
                        //'====== 下載次數規則分析 開始  ======
                        //if (String.IsNullOrEmpty(Convert.ToString(vds_allow_download_count)) == false)
                        if (vds_allow_download_count > 0)
                        {
                            //'自己上過的課 錄影檔
                            if (vds_video_type == 1)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND NOT EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number ) ";
                            }
                            //'別人上過的課 錄影檔
                            else if (vds_video_type == 2)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number AND session_type='RG' ) ";
                            }
                            //'大會堂 錄影檔
                            else if (vds_video_type == 3)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number AND session_type='SS' ) ";
                            }
                            //'不拘 錄影檔 當判斷條件為次數時, 不用下條件
                            else if (vds_video_type == 9)
                            {
                                str_condition1 = "";
                            }

                            str_sql_statement1 = " SELECT COUNT(*) AS used_count FROM sessionrecord_view_list WITH (NOLOCK) ";
                            str_sql_statement1 += " INNER JOIN SessionRecord_fileinfo WITH (NOLOCK) ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.sn ";
                            str_sql_statement1 += " WHERE ( sessionrecord_view_list.client_sn = @client_sn2 ) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.view_datetime >= @vds_valid_sdate2 ) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.view_datetime <= @vds_valid_edate2 ) ";
                            //str_sql_statement1 += " AND (sessionrecord_view_list.valid = 1)  AND ( isnull(sessionrecord_view_list.vds_video_type, '') <> '') AND (sessionrecord_view_list.vds_video_type = @video_type2) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.valid = 1 )  AND ( sessionrecord_view_list.vds_seq > '' ) ";
                            str_sql_statement1 += str_condition1;
                            user_SessionTable1 = ExcuteSql(str_sql_statement1, paraList);

                            // 例外錯誤
                            if (user_SessionTable1 == null)
                            {
                                return -9;
                            }

                            if (user_SessionTable1.Rows.Count > 0)
                            {
                                foreach (DataRow aRow1 in user_SessionTable1.Rows)
                                {
                                    used_count = (int)Convert.ToInt32(aRow1["used_count"]);   //已經使用的次數
                                }
                            }

                            if (String.IsNullOrEmpty(Convert.ToString(used_count)) == true)
                            {
                                used_count = 0;
                            }

                            //'檢查下載次數是否已達上限
                            if (used_count >= vds_allow_download_count)
                            {
                                ouput_value = -4;
                            }
                        }
                        //'====== 下載次數規則分析 結束  ======

                        //'====== 下載堂數規則分析 開始  ======
                        //if (String.IsNullOrEmpty(Convert.ToString(vds_allow_course_count)) == false)
                        else if (vds_allow_course_count > 0)
                        {

                            //'不拘 錄影檔 當判斷條件為次數時, 不用下條件
                            //'自己上過的課 錄影檔
                            if (vds_video_type == 1)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND NOT EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number ) ";
                            }
                            //'別人上過的課 錄影檔
                            else if (vds_video_type == 2)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number AND session_type='RG' ) ";
                            }
                            //'大會堂 錄影檔
                            else if (vds_video_type == 3)
                            {
                                paraList.Add(CreateSqlParameter("@video_type2", video_type));
                                str_condition1 = " AND ( sessionrecord_view_list.vds_seq = @video_type2 ) ";
                                str_condition1 += " AND EXISTS ( SELECT course FROM SessionRecord_selected_by_material WHERE session_sn = SessionRecord_fileinfo.session_record_number AND session_type='SS' ) ";
                            }
                            else if (vds_video_type == 9)
                            {
                                str_condition1 = "";
                                ArrayList paraList2 = new ArrayList();
                                //paraList2.Clear();
                                //錄影檔類型編號為10 for VIPABC
                                intSessionTypeSn = 10;
                                paraList2.Add(CreateSqlParameter("@client_sn3", client_sn));
                                paraList2.Add(CreateSqlParameter("@SessionTypeSn", intSessionTypeSn));
                                //讀出贈送錄影檔
                                str_sql_statement2 = " SELECT ISNULL(SUM(ContractChangeRecord.GiftPoint), 0) AS GiftPoint ";
                                str_sql_statement2 += " FROM ContractChangeRecord WITH (NOLOCK) INNER JOIN client_purchase WITH (NOLOCK) ON ContractChangeRecord.ContractSn = client_purchase.contract_sn ";
                                str_sql_statement2 += " WHERE ContractChangeRecord.valid = 1 AND client_purchase.client_sn = @client_sn3 AND ContractChangeRecord.SessionTypeSn = @SessionTypeSn ";
                                str_sql_statement2 += " AND ( CONVERT(VARCHAR, client_purchase.product_sdate, 111) <= CONVERT(VARCHAR, GETDATE(), 111) ) ";
                                str_sql_statement2 += " AND ( CONVERT(VARCHAR, client_purchase.product_edate, 111) >= CONVERT(VARCHAR, GETDATE(), 111) ) ";

                                user_SessionTable2 = ExcuteSql(str_sql_statement2, paraList2);

                                // 例外錯誤
                                if (user_SessionTable2 == null)
                                {
                                    return -9;
                                }

                                if (user_SessionTable2.Rows.Count > 0)
                                {
                                    foreach (DataRow aRow2 in user_SessionTable2.Rows)
                                    {
                                        gift_total_count = (float)Convert.ToDouble(aRow2["GiftPoint"]);   //贈送錄影檔
                                    }
                                }
                            }

                            str_sql_statement1 = " SELECT SUM(sessionrecord_view_list.view_cost) AS used_cost FROM sessionrecord_view_list WITH (NOLOCK) ";
                            str_sql_statement1 += " INNER JOIN SessionRecord_fileinfo WITH (NOLOCK) ON sessionrecord_view_list.view_session_recordfiles_sn = SessionRecord_fileinfo.sn ";
                            str_sql_statement1 += " WHERE ( sessionrecord_view_list.client_sn = @client_sn2 ) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.view_datetime >= @vds_valid_sdate2 ) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.view_datetime <= @vds_valid_edate2 ) ";
                            //str_sql_statement1 += " AND (sessionrecord_view_list.valid = 1)  AND ( isnull(sessionrecord_view_list.vds_seq, '') <> '') AND (sessionrecord_view_list.vds_seq = @video_type2) ";
                            str_sql_statement1 += " AND ( sessionrecord_view_list.valid = 1 )  AND ( sessionrecord_view_list.vds_seq > '' ) ";
                            str_sql_statement1 += str_condition1;

                            user_SessionTable1 = ExcuteSql(str_sql_statement1, paraList);

                            // 例外錯誤
                            if (user_SessionTable1 == null)
                            {
                                return -9;
                            }

                            if (user_SessionTable1.Rows.Count > 0)
                            {
                                foreach (DataRow aRow1 in user_SessionTable1.Rows)
                                {
                                    used_cost = (float)Convert.ToDouble(aRow1["used_cost"]);   //已經使用的次數
                                }
                            }

                            if (String.IsNullOrEmpty(Convert.ToString(used_cost)) == true)
                            {
                                used_cost = 0;
                            }

                            //'檢查下載堂數是否已達上限
                            //加入贈送錄影檔堂數
                            if (used_cost >= (vds_allow_course_count + gift_total_count))
                            {
                                ouput_value = -5;
                            }
                        }
                        //'====== 下載堂數規則分析 結束  ======
                    }
      
                    if (ouput_value == 0)
                    {
                        ouput_value = vds_subtract_points;
                    }
                }
                //'沒有設定規則
                else
                {
                    ouput_value = -2;
                }
            }

            //'顯示錯誤訊息
            if (ouput_value == -1)
            {
                //response.write "<script language='javascript1.2'>alert('表示SQL執行失敗');</script>"
                //'response.end
            }
            else if (ouput_value == -3)
            {
                //response.write "<script language='javascript1.2'>alert('表示傳入的ClientSn或VideoType是空值');</script>"
                //'response.end
            }
      	}
        catch (Exception ex)
        {
            //Console.WriteLine(ex.Message);
            ouput_value = -9;
       	}
        return ouput_value;
	}

    //工單:BD09050122、BD09050603 特殊扣點觀看錄影檔設定 raychien 20090511
    public string GetSeqFromDownloadVideoRule(int ClientSn, string VideoType)
    //'GetSeqFromDownloadVideoRule(ClientSn:Integer, VideoType:Integer): String
    //'函式描述: 傳回 根據為客戶設定特殊扣點觀看錄影檔原則 最新的序號
    //'Input:    ClientSn：   客戶編號
    //'          VideoType：  錄影檔類型                  [video_download_setting.vds_video_type]
    //'Ouput: Success: 回傳 Seq As String
    //'       Failed:  回傳 -1，表示SQL執行失敗
    //'                回傳 -2，表示沒有為此客戶設定任何扣點規則
    //'                回傳 -3，表示傳入的ClientSn或VideoType是空值
    //'                回傳 -9，例外錯誤
    {
        bool bo_error_happened;
        string ouput_value = "", str_sql_statement = "", video_type = "";
        int client_sn = 0;

        try
        {
            //'是否有錯誤發生
            bo_error_happened = false;

            //'回傳值
            ouput_value = "";

            client_sn = ClientSn;
            video_type = VideoType;

            if (String.IsNullOrEmpty(Convert.ToString(client_sn)) == true || String.IsNullOrEmpty(video_type) == true)
            {
                bo_error_happened = true;
                ouput_value = "-3";
            }

            if (bo_error_happened == false)
            {
                ArrayList paraList = new ArrayList();
                paraList.Add(CreateSqlParameter("@client_sn", client_sn));
                paraList.Add(CreateSqlParameter("@video_type", video_type));

                //'取出為客戶設定特殊扣點觀看錄影檔原則 最新的一筆
                str_sql_statement = "   SELECT TOP 1 vds_seq FROM video_download_setting  ";
                str_sql_statement += "  WHERE (vds_valid = 1)   ";
                str_sql_statement += "  AND (CONVERT(varchar, vds_valid_sdate, 111) <= CONVERT(varchar, getdate(), 111))     ";
                str_sql_statement += "  AND (CONVERT(varchar, vds_valid_edate, 111) >= CONVERT(varchar, getdate(), 111))     ";
                str_sql_statement += "  AND (clb_sn = @client_sn)    ";
                str_sql_statement += "  AND (vds_video_type = @video_type) ";
                str_sql_statement += "  ORDER BY vds_seq DESC    ";

                DataTable user_SessionTable = ExcuteSql(str_sql_statement, paraList);

                // 例外錯誤
                if (user_SessionTable == null)
                    return "-9";

                if (user_SessionTable.Rows.Count > 0)
                {
                    foreach (DataRow aRow in user_SessionTable.Rows)
                    {
                        ouput_value = Convert.ToString(aRow["vds_seq"].ToString());      // 序號
                    }
                }
                else
                {
                    ouput_value = "-2";
                }
            }
        }
        catch (Exception ex)
        {
            //Console.WriteLine(ex.Message);
            ouput_value = "-9";
        }
        return ouput_value;
    }


    public SqlParameter CreateSqlParameter(string paraName, object paraValue)
    {
        SqlParameter parameter = new SqlParameter(paraName, paraValue);
        return parameter;
    }

    public DataTable ExcuteSql(string sqlCommand, ArrayList sqlParameterList)
    {
        DataSet excute_Dataset = new DataSet();
        SqlConnection conn = null;
        SqlCommand cmd = null;
        DataTable dtb = null;
        try
        {
            dtb = new DataTable();
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sqlCommand, conn);
            if (sqlParameterList != null)
            {
                foreach (SqlParameter aSqlParameter in sqlParameterList)
                {
                    if (cmd.Parameters.IndexOf(aSqlParameter) == -1)
                    {
                        cmd.Parameters.Add(aSqlParameter);
                    }
                }
            }
            dtb.Load(cmd.ExecuteReader());

        }
        catch (Exception ex)
        {

            Console.WriteLine(ex.Message);
            return null;
        }
        finally
        {
            cmd.Parameters.Clear();
            if (cmd != null)
            {
                cmd.Dispose();
            }
            if (conn != null)
                conn.Close();
            cmd = null;
            conn = null;
        }
        return dtb;
    }

    public bool SQLExecuteNonQuery(string sqlCommand, ArrayList sqlParameterList)
    {
        bool isok = false;
        SqlConnection conn = null;
        SqlCommand cmd = null;
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sqlCommand, conn);
            if (sqlParameterList != null)
            {
                foreach (SqlParameter aSqlParameter in sqlParameterList)
                {
                    if (cmd.Parameters.IndexOf(aSqlParameter) == -1)
                    {
                        cmd.Parameters.Add(aSqlParameter);
                    }
                }
            }
            if (cmd.ExecuteNonQuery() > 0)
            {
                isok = true;
            }
            else
            {
                isok = false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            cmd.Parameters.Clear();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }

        return isok;
    }

    public VideoInfo createVideoInfoByEncode(string p_strEncode)
    {
        VideoInfo aVideoInfo = new VideoInfo();
        string strSql = " SELECT dbo.material.ltitle, dbo.con_basic.basic_fname + N' ' + dbo.con_basic.basic_lname AS con_name, dbo.client_basic.fname + N' ' + dbo.client_basic.lname AS client_name, " +
                               " dbo.client_attend_list.attend_level, DATEADD(hh, dbo.client_attend_list.attend_sestime, dbo.client_attend_list.attend_date) AS attend_date, dbo.SessionRecord_fileinfo.File_Full_name, " +
                               " dbo.client_attend_list.attend_consultant, dbo.SessionRecord_fileinfo.SN AS file_sn, dbo.client_attend_list.session_sn, dbo.client_attend_list.client_sn, dbo.SessionRecord_fileinfo.File_path, dbo.material.course " +
                               " FROM dbo.SessionRecord_fileinfo INNER JOIN " +
                               " dbo.client_attend_list ON dbo.SessionRecord_fileinfo.Session_Record_Number = dbo.client_attend_list.session_sn INNER JOIN " +
                               " dbo.con_basic ON dbo.client_attend_list.attend_consultant = dbo.con_basic.con_sn INNER JOIN " +
                               " dbo.material ON dbo.client_attend_list.attend_mtl_1 = dbo.material.course INNER JOIN " +
                               " dbo.client_basic ON dbo.client_attend_list.client_sn = dbo.client_basic.sn" +
                               " WHERE (dbo.SessionRecord_fileinfo.encodestr = @ecnodeStr) ";

        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {

            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(strSql, conn);
            cmd.Parameters.Add("@ecnodeStr", SqlDbType.VarChar);
            cmd.Parameters["@ecnodeStr"].Value = p_strEncode;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return aVideoInfo;
            while (rs.Read())
            {
                aVideoInfo.Consultant = rs["attend_consultant"].ToString();
                aVideoInfo.SessionStudent.Add(rs["client_sn"].ToString());
                aVideoInfo.NameOfVideo = rs["ltitle"].ToString();
                aVideoInfo.ConsultantName = rs["con_name"].ToString();
                aVideoInfo.SessionLevel = Convert.ToInt32(rs["attend_level"]);
                aVideoInfo.File_Date = Convert.ToDateTime(rs["attend_date"]);
                aVideoInfo.File_FullName = rs["File_Full_name"].ToString();
                aVideoInfo.File_SN = rs["file_sn"].ToString();
                aVideoInfo.SessionSN = rs["session_sn"].ToString();
                aVideoInfo.IsOwnSession = true;
                aVideoInfo.File_Path = getServerPath(rs["file_path"].ToString());
                aVideoInfo.Course = rs["course"].ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return aVideoInfo;
    }

    private string getServerPath(string str_server_sn)
    {
        string server_id = "";
        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        try
        {
            string sql = " SELECT path_text FROM cfg_sessionfile_path(nolock) WHERE (sn =" + str_server_sn + ") ";
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(sql, conn);
            rs = cmd.ExecuteReader();
            while (rs.Read())
            {
                server_id = rs["path_text"].ToString().Replace("svr", "").Replace(@"/", "");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return server_id;
    }

    public String getClientName(String p_client_sn)
    {
        string strSql = " SELECT fname+N' '+lname as client_full_name " +
                               " FROM dbo.client_basic " +
                               " WHERE sn = @client_sn";

        SqlConnection conn = null;
        SqlCommand cmd = null;
        SqlDataReader rs = null;
        String strClientName = "";
        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["muchnewdb"].ToString());
            conn.Open();
            cmd = new SqlCommand(strSql, conn);
            cmd.Parameters.Add("@client_sn", SqlDbType.Int);
            cmd.Parameters["@client_sn"].Value = p_client_sn;
            rs = cmd.ExecuteReader();
            if (rs == null)
                return null;
            while (rs.Read())
            {
                strClientName = rs["client_full_name"].ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (rs != null)
                rs.Close();
            if (cmd != null)
                cmd.Dispose();
            if (conn != null)
                conn.Dispose();
        }
        return strClientName;
    }

    public string GetHtml(string Url)
    {
        WebRequest WRq = WebRequest.Create(Url);
        WebResponse WRs = WRq.GetResponse();
        Stream s = WRs.GetResponseStream();
        StreamReader sr = new StreamReader(s, System.Text.Encoding.UTF8);
        string res = sr.ReadToEnd();
        sr.Close();
        sr.Dispose();
        s.Close();
        s.Dispose();
        WRs.Close();
        return res;
    }

    public int getVideoLength() {
        return intSelectLength;
    }
    private void setVideoLength(int dataLength)
    {
        intSelectLength = dataLength;
    }
}